# -*- coding: utf-8 -*-
# Rhino 8 for macOS — Generate ClippingDrawings per DECK_* section and export DWG per deck
# Pipeline per deck:
#   generate_drawing(sectionName) -> export_sublayers_dwg(layer) -> cleanup_drawing(layer)
#   Note: We also delete the temporary layers created by ClippingDrawings after export.
#
# Notes:
# - Uses ClippingDrawings with:
#   Angle=0, Projection=Parallel, AddSilhouette=Yes, ShowHatch=Yes, ShowSolid=Yes,
#   AddBackground=Yes, ShowLabel=No. (PrintWidth/DisplayColor set to “By Input Object” if your build supports those tokens.)
# - Placement: moves the generated drawing so its center is 3× model width to the RIGHT of the model’s bbox center.
# - Exports ONLY the drawing layer (including sublayers) to ~/Desktop/<sectionName>-Export.dwg
# - No layouts, no clipboard, no ChangeSpace.

import rhinoscriptsyntax as rs
import Rhino
import re
import os
import time
import logging
import tempfile

# ---------- logging ----------
# We initialize a module-level logger that writes verbose diagnostics to a log file
# and a concise stream to the console (Rhino command history). The file handler
# captures DEBUG+ while the console emits INFO+ for readability.
logger = logging.getLogger("clipping_export")


def _default_log_path():
    """
    Determine a writable log file path.
    Preference order:
      1) ~/Desktop/RhinoClippingExport.log
      2) current working directory ./RhinoClippingExport.log
      3) system temporary directory
    """
    try:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        if os.path.isdir(desktop):
            return os.path.join(desktop, "RhinoClippingExport.log")
    except Exception:
        # Fall through to next option
        pass
    try:
        cwd = os.getcwd()
        return os.path.join(cwd, "RhinoClippingExport.log")
    except Exception:
        pass
    # Last resort: temp directory
    return os.path.join(tempfile.gettempdir(), "RhinoClippingExport.log")


def _setup_logging():
    """
    Configure logging only once per interpreter session.
    - FileHandler captures DEBUG and above (full fidelity)
    - StreamHandler captures INFO and above (human friendly)
    """
    if getattr(_setup_logging, "_configured", False):
        return

    log_path = _default_log_path()
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        fmt="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # File handler for full diagnostics
    try:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except Exception:
        # If file handler fails, at least keep console logging
        logger.debug(
            "FileHandler init failed; continuing with console-only.", exc_info=True)

    # Console handler for Rhino command history
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)

    _setup_logging._configured = True
    logger.debug("Logging initialized. Log file: %s", log_path)


# Initialize logging on import so any functions used standalone still log.
_setup_logging()

# ---------- helpers ----------


def _activate_model_view():
    """Ensure we’re in a model-space viewport (not a layout)."""
    logger.debug("Activating a model-space viewport.")
    doc = Rhino.RhinoDoc.ActiveDoc
    for v in doc.Views:
        if not isinstance(v, Rhino.Display.RhinoPageView):
            doc.Views.ActiveView = v
            try:
                rs.CurrentView(v.ActiveViewport.Name)
            except Exception:
                logger.debug("Failed to set CurrentView to %s", getattr(
                    v.ActiveViewport, "Name", "?"), exc_info=True)
            rs.Redraw()
            logger.debug("Activated model-space view: %s",
                         getattr(v.ActiveViewport, "Name", "?"))
            return True
    created = rs.Command('_-NewViewport _Top _Enter', echo=False)
    logger.debug("New viewport created via command: %s", created)
    return created


def _restore_named_view_if_exists(name):
    """Restore named view into current model viewport if it exists."""
    logger.debug("Attempting to restore named view: %s", name)
    doc = Rhino.RhinoDoc.ActiveDoc
    idx = doc.NamedViews.FindByName(name)
    if idx >= 0:
        vp = doc.Views.ActiveView.ActiveViewport
        try:
            doc.NamedViews.Restore(idx, vp, True)
            logger.debug("Named view restored: %s", name)
        except Exception:
            logger.debug("Failed to restore named view: %s",
                         name, exc_info=True)


def _find_clipping_plane_by_name(name):
    """Return the first ClipPlaneObject whose name == name."""
    doc = Rhino.RhinoDoc.ActiveDoc
    for obj in doc.Objects.GetObjectList(Rhino.DocObjects.ObjectType.ClipPlane):
        try:
            if (obj.Attributes.Name or "") == name:
                logger.debug("Found clipping plane with name: %s", name)
                return obj
        except Exception:
            logger.debug("Error while scanning clip planes", exc_info=True)
    return None


def _model_bbox():
    """Bounding box of all model-space geometry."""
    logger.debug("Computing model-space bounding box.")
    doc = Rhino.RhinoDoc.ActiveDoc
    bb = None
    for o in doc.Objects:
        try:
            if o.IsDeleted:
                continue
            if o.Attributes.Space != Rhino.DocObjects.ActiveSpace.ModelSpace:
                continue
            ob = o.Geometry.GetBoundingBox(True)
            if not ob.IsValid:
                continue
            bb = ob if bb is None else Rhino.Geometry.BoundingBox.Union(bb, ob)
        except Exception:
            logger.debug(
                "Error computing object bbox; skipping one object.", exc_info=True)
    logger.debug("Model bbox computed: %s", bb)
    return bb


def _move_ids_right_of_model(ids, factor=3.0):
    """Move ids so their center is factor×model width to the right of model center."""
    if not ids:
        return
    logger.debug(
        "Moving %d ids to the right of model by factor %.2f", len(ids), factor)
    mbb = _model_bbox()
    if not mbb or not mbb.IsValid:
        return
    mmin, mmax = mbb.Min, mbb.Max
    mw = max(0.0, mmax.X - mmin.X)
    mcx = 0.5*(mmin.X + mmax.X)
    mcy = 0.5*(mmin.Y + mmax.Y)

    bb = rs.BoundingBox(ids)
    if not bb or len(bb) < 2:
        return
    minx = min(p.X for p in bb)
    maxx = max(p.X for p in bb)
    miny = min(p.Y for p in bb)
    maxy = max(p.Y for p in bb)
    dcx = 0.5*(minx + maxx)
    dcy = 0.5*(miny + maxy)

    dx = (mcx + factor*mw) - dcx
    dy = mcy - dcy
    try:
        rs.MoveObjects(ids, (dx, dy, 0.0))
        logger.debug("Moved objects by dx=%.3f, dy=%.3f", dx, dy)
    except Exception:
        logger.debug("Failed to move objects.", exc_info=True)


def _ensure_layer(name, color=None, parent=None):
    """Create layer (optionally under parent) and return full layer path."""
    full = name
    if parent:
        full = parent + "::" + name
    if not rs.IsLayer(full):
        try:
            rs.AddLayer(full, color=color)
            logger.debug("Created layer: %s", full)
        except Exception:
            logger.debug("Failed to create layer: %s", full, exc_info=True)
    return full


def _diff_new_objects(before_set):
    after = set(rs.AllObjects() or [])
    diff = list(after - before_set)
    logger.debug("Detected %d new objects.", len(diff))
    return diff


def _diff_new_layers(before_set):
    """
    Return list of new layer names created since 'before_set' snapshot.
    """
    after = set(rs.LayerNames() or [])
    diff = [ln for ln in (after - before_set)]
    logger.debug("Detected %d new layers.", len(diff))
    return diff


def _objs_on_layer_and_children(layer_name):
    """Return object ids on layer and sublayers."""
    ids = []
    all_layers = rs.LayerNames() or []
    targets = []
    for ln in all_layers:
        if ln == layer_name or ln.startswith(layer_name + "::"):
            targets.append(ln)
    if not targets and rs.IsLayer(layer_name):
        targets = [layer_name]
    for ln in targets:
        try:
            ids.extend(rs.ObjectsByLayer(ln, True) or [])
        except Exception:
            logger.debug("Failed to collect objects for layer: %s",
                         ln, exc_info=True)
    # unique
    uniq = []
    seen = set()
    for i in ids:
        if i not in seen:
            uniq.append(i)
            seen.add(i)
    logger.debug("Collected %d objects across %d layers (including sublayers).", len(
        uniq), len(targets))
    return uniq, targets


def _unlock_layers(layer_names):
    logger.debug("Unlocking %d layers.", len(layer_names))
    for ln in layer_names:
        try:
            if rs.IsLayer(ln) and rs.IsLayerLocked(ln):
                rs.UnlockLayer(ln)
                logger.debug("Unlocked layer: %s", ln)
        except Exception:
            logger.debug("Failed to unlock layer: %s", ln, exc_info=True)


def _lock_layers(layer_names):
    logger.debug("Locking %d layers.", len(layer_names))
    for ln in layer_names:
        try:
            if rs.IsLayer(ln) and not rs.IsLayerLocked(ln):
                rs.LockLayer(ln)
                logger.debug("Locked layer: %s", ln)
        except Exception:
            logger.debug("Failed to lock layer: %s", ln, exc_info=True)


def _delete_layer_tree(layer_name):
    """Delete layer and all its children."""
    all_layers = sorted(rs.LayerNames() or [],
                        key=lambda s: len(s.split("::")), reverse=True)
    # delete children first
    logger.info("Deleting layer tree: %s", layer_name)
    for ln in all_layers:
        if ln == layer_name or ln.startswith(layer_name + "::"):
            try:
                # delete objects on the layer
                objs = rs.ObjectsByLayer(ln, True) or []
                if objs:
                    for oid in objs:
                        try:
                            rs.DeleteObject(oid)
                        except Exception:
                            logger.debug(
                                "Failed to delete object on layer: %s", ln, exc_info=True)
                # delete layer
                if rs.IsLayer(ln):
                    try:
                        rs.DeleteLayer(ln)
                        logger.debug("Deleted layer: %s", ln)
                    except Exception:
                        logger.debug("Failed to delete layer: %s",
                                     ln, exc_info=True)
            except Exception:
                logger.debug(
                    "Error while deleting within layer tree for: %s", ln, exc_info=True)

# ---------- file/system helpers ----------


def _wait_for_file(path, timeout_seconds=20.0, poll_seconds=0.25):
    """
    Wait until a file at 'path' exists, is non-empty, and its size has stabilized.
    This helps with async/slow writer plugins so we don't prematurely fail.
    Returns True if the file appears stable before timeout; otherwise False.
    """
    logger.debug("Waiting for file to stabilize: %s (timeout=%.1fs)",
                 path, timeout_seconds)
    start = time.time()
    last_size = -1
    stable_count = 0
    while (time.time() - start) < timeout_seconds:
        try:
            if os.path.exists(path):
                size = os.path.getsize(path)
                if size > 0:
                    if size == last_size:
                        stable_count += 1
                        # consider stable after two consecutive identical sizes
                        if stable_count >= 2:
                            logger.debug(
                                "File is stable (size=%d): %s", size, path)
                            return True
                    else:
                        stable_count = 0
                        last_size = size
        except Exception:
            logger.debug("Error while probing file status for: %s",
                         path, exc_info=True)
        time.sleep(poll_seconds)
    logger.debug("Timed out waiting for file: %s", path)
    return False

# Resolve DWG output path helper


def _resolve_export_outpath(sectionName, output_dir=None):
    """
    Compute the DWG output path for a given section name.
    - If output_dir is provided and valid, place file there
    - Otherwise primary: user's Desktop; fallback: current working directory
    Filename: <sectionName>-Export.dwg
    """
    try:
        if output_dir and os.path.isdir(output_dir):
            base = output_dir
        else:
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            base = desktop if os.path.isdir(desktop) else os.getcwd()
        return os.path.join(base, "{}-Export.dwg".format(sectionName))
    except Exception:
        # Last resort: ensure we return something writable with a safe filename
        safe_name = "{}-Export.dwg".format(
            re.sub(r"[^\w\-\.]+", "_", sectionName))
        return os.path.join(output_dir or tempfile.gettempdir(), safe_name)

# ---------- required API ----------


def generate_drawing(sectionName):
    """
    1) Select clipping plane named sectionName and run ClippingDrawings with:
       Angle=0, Projection=Parallel, AddSilhouette=Yes, ShowHatch=Yes, ShowSolid=Yes, AddBackground=Yes, ShowLabel=No.
       (Tries to set PrintWidth/DisplayColor to ByInputObject if available.)
    2) Move the result 3× model width to the right of model center.
    3) Create/ensure a layer 'DRAWING_<sectionName>' and move all result objects there.
    4) Return (drawing_layer_name, created_layers_from_clipping_drawings).
       'created_layers_from_clipping_drawings' are temporary layers produced by the
       ClippingDrawings command that we will delete after export to avoid layer bloat.
    """
    logger.info("Generating drawing for section: %s", sectionName)
    if not _activate_model_view():
        logger.error("No model-space view available.")
        raise Exception("No model-space view available")

    # If there is a matching named view, restore it
    _restore_named_view_if_exists(sectionName)

    # Find the clipping plane object
    cp = _find_clipping_plane_by_name(sectionName)
    if cp is None:
        logger.error("Clipping plane not found: %s", sectionName)
        raise Exception("Clipping plane not found: {}".format(sectionName))

    # Snapshot layers before running ClippingDrawings so we can detect which layers it creates
    before_layers = set(rs.LayerNames() or [])

    # Run ClippingDrawings with requested options
    before = set(rs.AllObjects() or [])
    rs.UnselectAllObjects()
    try:
        rs.SelectObject(cp.Id)
        logger.debug("Selected clipping plane id: %s", getattr(cp, "Id", None))
    except Exception:
        logger.debug("Failed to select clipping plane id for: %s",
                     sectionName, exc_info=True)

    # Important: some mac builds use AddBackground (not ShowBackground)
    # PrintWidth/DisplayColor tokens may be unavailable; include them if accepted, otherwise Rhino ignores.
    cmd = (
        '-_ClippingDrawings '
        'Angle 0 '
        'Projection=_Parallel '
        'AddSilhouette=_Yes '
        'ShowHatch=_Yes '
        'ShowSolid=_Yes '
        'AddBackground=_Yes '
        'ShowLabel=_No '
        'PrintWidth=_ByInputObject '
        'DisplayColor=_ByInputObject '
        '_Enter'
    )
    logger.debug("Executing command: %s", cmd)
    rs.Command(cmd, echo=False)

    # Determine which layers were created by ClippingDrawings (before we create our own DRAWING_* layer)
    created_layers = _diff_new_layers(before_layers)

    new_ids = _diff_new_objects(before)
    if not new_ids:
        logger.error(
            "ClippingDrawings produced no geometry for %s", sectionName)
        raise Exception(
            "ClippingDrawings produced no geometry for {}".format(sectionName))

    # Place drawing right of model
    _move_ids_right_of_model(new_ids, factor=3.0)

    # Create/ensure target layer and move results there
    drawing_layer = "DRAWING_{}".format(sectionName)
    _ensure_layer(drawing_layer)
    for oid in new_ids:
        try:
            rs.ObjectLayer(oid, drawing_layer)
        except Exception:
            logger.debug("Failed to move object to layer: %s",
                         drawing_layer, exc_info=True)

    logger.info("Generated drawing on layer: %s (objects: %d)",
                drawing_layer, len(new_ids))
    return drawing_layer, created_layers


def export_sublayers_dwg(drawingLayerName, output_dir=None, allow_overwrite=True):
    """
    Unlock layer + all sublayers, select all their objects, and export ONLY those to DWG.
    Filename: <sectionName>-Export.dwg (sectionName is the part after 'DRAWING_').
    Then re-lock the layers.

    Args:
        drawingLayerName: Name of the drawing layer tree, must start with 'DRAWING_'.
        output_dir: Optional folder where the DWG should be written.
        allow_overwrite: If False and file exists, skip writing and return existing path.
    """
    # Resolve export filename
    if not drawingLayerName.startswith("DRAWING_"):
        logger.error("Unexpected drawing layer name: %s", drawingLayerName)
        raise Exception(
            "Unexpected drawing layer name: {}".format(drawingLayerName))
    sectionName = drawingLayerName[len("DRAWING_"):]
    outpath = _resolve_export_outpath(sectionName, output_dir=output_dir)
    logger.info("Exporting DWG for layer tree '%s' -> %s",
                drawingLayerName, outpath)

    # Respect overwrite policy early
    try:
        if (not allow_overwrite) and os.path.exists(outpath) and os.path.getsize(outpath) > 0:
            logger.info("Skipping export (preserving existing): %s", outpath)
            return outpath
    except Exception:
        logger.debug("Overwrite check failed for: %s", outpath, exc_info=True)

    # Collect layers and objects
    ids, layers = _objs_on_layer_and_children(drawingLayerName)
    if not ids:
        logger.error("No objects on layer tree: %s", drawingLayerName)
        raise Exception(
            "No objects on layer tree: {}".format(drawingLayerName))

    # Unlock layers, select ids
    _unlock_layers(layers)
    rs.UnselectAllObjects()
    try:
        rs.SelectObjects(ids)
    except Exception:
        logger.debug("Failed to select objects for export.", exc_info=True)

    # Remove stale file only when overwrite is allowed
    if allow_overwrite:
        try:
            if os.path.exists(outpath):
                os.remove(outpath)
                logger.debug(
                    "Removed existing file prior to export: %s", outpath)
        except Exception:
            logger.debug(
                "Failed to remove existing export file (continuing): %s", outpath, exc_info=True)

    wrote = False

    # First attempt: RhinoCommon WriteFile with "selected only", suppressing UI.
    # This avoids command-line option variance across mac builds.
    try:
        logger.debug(
            "Attempting RhinoCommon WriteFile (selected-only) to: %s", outpath)
        doc = Rhino.RhinoDoc.ActiveDoc
        opts = Rhino.FileIO.FileWriteOptions()
        opts.WriteSelectedObjectsOnly = True
        opts.SuppressAllInput = True
        # Some exporters honor geometry-only when selection is present; we rely on 'selected only'
        ok = doc.WriteFile(outpath, opts)
        logger.debug("WriteFile returned: %s", ok)
        if _wait_for_file(outpath, timeout_seconds=30.0):
            wrote = True
            logger.info("DWG export succeeded via RhinoCommon: %s", outpath)
    except Exception:
        logger.debug(
            "RhinoCommon WriteFile failed, will try command-based export.", exc_info=True)

    # Fallback: Command-driven export (mac builds differ in prompts).
    if not wrote:
        export_cmds = [
            u'-_Export "{}" _Enter'.format(outpath),
            u'-_Export "{}" _Enter _Enter'.format(outpath),
            u'-_Export {} _Enter'.format(outpath),
            u'-_Export {} _Enter _Enter'.format(outpath),
        ]
        for cmd in export_cmds:
            logger.debug("Attempting export via command: %s", cmd)
            rs.Command(cmd, echo=False)
            # Wait up to 30s for the exporter plugin to finish writing
            if _wait_for_file(outpath, timeout_seconds=30.0):
                wrote = True
                logger.info("DWG export succeeded: %s", outpath)
                break

    # Re-lock layers regardless
    _lock_layers(layers)
    rs.UnselectAllObjects()

    if not wrote:
        logger.error("DWG export failed: %s", outpath)
        raise Exception("DWG export failed: {}".format(outpath))

    return outpath


def cleanup_drawing(drawingLayerName):
    """Delete the drawing layer (and all sublayers/objects) created for this deck."""
    logger.info("Cleaning up drawing layer tree: %s", drawingLayerName)
    _delete_layer_tree(drawingLayerName)


def export_deck(sectionName, output_dir=None, allow_overwrite=True):
    """
    Orchestrates a single deck export:
      1) generate_drawing(sectionName) -> (drawingLayerName, createdTempLayers)
      2) export_sublayers_dwg(drawingLayerName)
      3) delete createdTempLayers (from ClippingDrawings)
      4) cleanup_drawing(drawingLayerName)
      5) If DWG already exists before starting and allow_overwrite is False, skip.

    Args:
        sectionName: The clipping plane name to export.
        output_dir: Optional destination folder for the DWG.
        allow_overwrite: If False and the output file exists, skip this export.
    """
    logger.info("Starting deck export for: %s", sectionName)
    # Skip if output already exists to avoid re-exporting the same deck
    try:
        preexisting_out = _resolve_export_outpath(
            sectionName, output_dir=output_dir)
        if (not allow_overwrite) and os.path.exists(preexisting_out) and os.path.getsize(preexisting_out) > 0:
            logger.info("Skipping export for %s; DWG already exists: %s",
                        sectionName, preexisting_out)
            return preexisting_out
    except Exception:
        logger.debug(
            "Pre-export existence check failed; will proceed with export.", exc_info=True)

    drawing_layer = None
    created_temp_layers = []
    out = None
    try:
        drawing_layer, created_temp_layers = generate_drawing(sectionName)
        out = export_sublayers_dwg(
            drawing_layer, output_dir=output_dir, allow_overwrite=allow_overwrite)
        return out
    finally:
        # Always delete the temporary layers created by ClippingDrawings to avoid Rhino layer bloat
        if created_temp_layers:
            logger.info("Deleting %d temporary layers created by ClippingDrawings.", len(
                created_temp_layers))
            for ln in created_temp_layers:
                try:
                    if ln and rs.IsLayer(ln):
                        _delete_layer_tree(ln)
                except Exception:
                    logger.debug(
                        "Failed to delete temporary layer tree: %s", ln, exc_info=True)
        # Remove the generated drawing layer tree as before
        if drawing_layer:
            cleanup_drawing(drawing_layer)
        if out:
            logger.info("Export complete for %s -> %s", sectionName, out)
        else:
            logger.info(
                "Export finished with no output path for %s (likely failed earlier).", sectionName)

# ---------- discovery and interactive prompts ----------


def _all_section_names_with_prefix(prefix):
    """
    Collect names of all clipping planes whose name starts with the given prefix.
    Names are returned in case-insensitive sorted order for stability.
    """
    names = []
    doc = Rhino.RhinoDoc.ActiveDoc
    for obj in doc.Objects.GetObjectList(Rhino.DocObjects.ObjectType.ClipPlane):
        nm = obj.Attributes.Name or ""
        if nm.startswith(prefix):
            names.append(nm)
    names.sort(key=lambda s: s.lower())
    logger.info("Discovered %d sections with prefix '%s'.", len(names), prefix)
    return names


def _force_parallel_projection():
    # keep projection parallel in current view to match spec
    try:
        v = rs.CurrentView()
        if v:
            rs.ViewProjection(v, 1)  # 1 = parallel
            logger.debug("Forced parallel projection on view: %s", v)
    except Exception:
        logger.debug("Failed to force parallel projection.", exc_info=True)


def _prompt_prefix(default_prefix="FLOOR_"):
    """
    Prompt the user for the clipping section prefix.
    Defaults to 'FLOOR_' if user accepts default or submits empty.
    Returns the prefix string, or None if cancelled.
    """
    try:
        result = rs.GetString("Enter section prefix", default_prefix)
    except Exception:
        result = default_prefix
    if result is None:
        return None
    result = result.strip()
    return result if result else default_prefix


def _confirm_sections_list(sections):
    """
    Show the list of matched sections and ask the user to confirm to continue.
    Returns True to continue, False to abort.
    """
    if not sections:
        logger.warning("No sections matched the given prefix.")
        return False
    logger.info("Matched sections (%d): %s",
                len(sections), ", ".join(sections))
    try:
        resp = rs.GetString("Continue with these sections? (Y/N) [Y]", "Y")
    except Exception:
        resp = "Y"
    if resp is None:
        return False
    resp = (resp or "Y").strip().lower()
    return resp.startswith("y")


def _prompt_export_folder():
    """
    Prompt the user to select the export folder via UI; fallback to Desktop on text input.
    Returns folder path or None if cancelled.
    """
    try:
        folder = rs.BrowseForFolder(
            message="Select export folder for DWG files")
        if folder and os.path.isdir(folder):
            logger.info("Selected export folder: %s", folder)
            return folder
    except Exception:
        logger.debug(
            "BrowseForFolder failed, falling back to string prompt.", exc_info=True)
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    prompt = "Enter export folder path [{}]: ".format(desktop)
    try:
        result = rs.GetString(prompt, desktop)
    except Exception:
        result = desktop
    if result is None:
        return None
    result = os.path.expanduser((result or desktop).strip())
    if not os.path.isdir(result):
        logger.error("Invalid export folder: %s", result)
        return None
    logger.info("Selected export folder: %s", result)
    return result


def _select_existing_to_overwrite(section_to_path):
    """
    Given a mapping of section -> outpath, show which files already exist and
    let the user select which of them to overwrite.
    Returns a set of section names that the user chose to overwrite.
    If no files exist, returns an empty set (caller can skip selection).
    """
    existing = [(sec, p) for sec, p in section_to_path.items()
                if os.path.exists(p) and os.path.getsize(p) > 0]
    if not existing:
        return set()

    rows = []
    for sec, p in existing:
        rows.append([sec, p])

    selected_secs = set()
    # Try multi-selection dialog if available
    try:
        selection = rs.MultiListBox(rows, ["Section", "Existing File"],
                                    "Select existing files to OVERWRITE",
                                    "Overwrite Existing Exports")
        if selection:
            for row in selection:
                if row and len(row) > 0:
                    selected_secs.add(row[0])
            return selected_secs
    except Exception:
        logger.debug(
            "MultiListBox not available; falling back to typed selection.", exc_info=True)

    # Fallback: display and accept comma-separated indices
    logger.info("Existing exports:")
    for idx, (sec, p) in enumerate(existing, 1):
        logger.info("%d) %s -> %s", idx, sec, p)
    try:
        raw = rs.GetString(
            "Enter numbers to overwrite (comma-separated), or 'none'/'all' [none]", "none")
    except Exception:
        raw = "none"
    if raw is None:
        return set()
    raw = (raw or "none").strip().lower()
    if raw == "all":
        return {sec for sec, _ in existing}
    if raw == "none" or raw == "":
        return set()
    try:
        indices = [int(x.strip())
                   for x in raw.split(",") if x.strip().isdigit()]
        for i in indices:
            if 1 <= i <= len(existing):
                selected_secs.add(existing[i-1][0])
    except Exception:
        logger.debug(
            "Failed to parse selection; defaulting to none.", exc_info=True)
    return selected_secs


def main():
    """
    Interactive export workflow:
      1) Ask for section prefix (default 'FLOOR_')
      2) Show matched sections and confirm to continue
      3) Ask for export folder (file picker dialog)
      4) Show which export files already exist and let the user select which to overwrite
         (if none exist, this step is skipped)
      5) Export the selected sections
    """
    logger.info("Starting interactive export.")
    if not _activate_model_view():
        logger.error("No model view. Abort.")
        return
    _force_parallel_projection()

    # 1) Prefix prompt
    prefix = _prompt_prefix(default_prefix="FLOOR_")
    if prefix is None:
        logger.warning("Prefix prompt cancelled.")
        return

    # 2) Discover and confirm sections
    sections = _all_section_names_with_prefix(prefix)
    if not sections:
        logger.warning("No clipping sections starting with %s.", prefix)
        return
    if not _confirm_sections_list(sections):
        logger.info("User declined to continue with matched sections.")
        return

    # 3) Export folder
    export_folder = _prompt_export_folder()
    if export_folder is None:
        logger.error("Export folder selection cancelled. Aborting.")
        return

    # Map each section to its intended outpath
    section_to_outpath = {}
    for sec in sections:
        section_to_outpath[sec] = _resolve_export_outpath(
            sec, output_dir=export_folder)

    # 4) Existing file selection
    overwrite_existing_secs = _select_existing_to_overwrite(section_to_outpath)

    # Determine final set of sections to export:
    # - Always include those without an existing file
    # - Include those user selected to overwrite
    to_export = []
    for sec, outp in section_to_outpath.items():
        exists = False
        try:
            exists = os.path.exists(outp) and os.path.getsize(outp) > 0
        except Exception:
            exists = os.path.exists(outp)
        if (not exists) or (sec in overwrite_existing_secs):
            to_export.append(sec)
    if not to_export:
        logger.info("Nothing to export based on current selections.")
        return

    # 5) Export
    for sec in to_export:
        try:
            allow_overwrite = sec in overwrite_existing_secs
            export_deck(sec, output_dir=export_folder,
                        allow_overwrite=allow_overwrite)
        except Exception as e:
            logger.exception("FAIL %s -> %s", sec, e)


main()
