# -*- coding: utf-8 -*-
# Rhino 8 for macOS — Assemble layouts from exported DWGs and export vector PDF
#
# Behavior:
# - Discovers DWG files (by default: ~/Desktop/*-Export.dwg) or accepts explicit paths.
# - Optionally duplicates a master layout (if present) as the base for each page.
# - Creates one page layout per DWG and inserts a single detail viewport.
# - Sets the detail scale so that 1 mm on paper equals 0.2 m in the imported DWG by default.
# - Attempts to export a vector PDF of all layouts.
#
# Notes:
# - This script is designed to run in a mostly empty Rhino file. If a master
#   layout exists, its page contents will be copied to every generated page
#   by duplicating the master layout as a template.
# - Scaling can be adjusted via arguments to main().

import rhinoscriptsyntax as rs
import Rhino
import os
import glob
import time
import logging
import tempfile

# ---------- logging ----------
logger = logging.getLogger("assemble_layouts")


def _default_log_path():
    """
    Determine a writable log file path.
    Preference order:
      1) ~/Desktop/LayoutAssembly.log
      2) current working directory ./LayoutAssembly.log
      3) system temporary directory
    """
    try:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        if os.path.isdir(desktop):
            return os.path.join(desktop, "LayoutAssembly.log")
    except Exception:
        pass
    try:
        return os.path.join(os.getcwd(), "LayoutAssembly.log")
    except Exception:
        pass
    return os.path.join(tempfile.gettempdir(), "LayoutAssembly.log")


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
    try:
        fh = logging.FileHandler(log_path, encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except Exception:
        logger.debug(
            "FileHandler init failed; continuing with console-only.", exc_info=True)
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    _setup_logging._configured = True
    logger.debug("Logging initialized. Log file: %s", log_path)


_setup_logging()

# ---------- user input helpers ----------


def _get_page_format_dimensions(format_name, landscape=False):
    """
    Get page width and height in millimeters for a given format name.

    Args:
        format_name: Page format name (A4, A3, A2, A1) - case insensitive
        landscape: If True, swap width and height for landscape orientation

    Returns:
        tuple: (width_mm, height_mm) or None if format not recognized
    """
    formats = {
        'A4': (210.0, 297.0),
        'A3': (297.0, 420.0),
        'A2': (420.0, 594.0),
        'A1': (594.0, 841.0),
    }
    dimensions = formats.get(format_name.upper().strip())
    if dimensions and landscape:
        # Swap width and height for landscape
        return (dimensions[1], dimensions[0])
    return dimensions


def _prompt_orientation():
    """
    Prompt user to select page orientation (Portrait or Landscape).

    Returns:
        bool: True for landscape, False for portrait, or None if cancelled
    """
    prompt = "Select orientation (Portrait/Landscape) [Portrait]: "
    default = "Portrait"
    result = rs.GetString(prompt, default)

    if result is None:
        return None

    result = result.strip()
    if not result:
        result = default

    result_lower = result.lower()
    if result_lower.startswith('l'):
        logger.info("Selected orientation: Landscape")
        return True
    elif result_lower.startswith('p'):
        logger.info("Selected orientation: Portrait")
        return False
    else:
        logger.warning(
            "Invalid orientation '%s', using default Portrait", result)
        return False


def _prompt_page_format():
    """
    Prompt user to select a page format (A4, A3, A2, A1) and orientation.

    Returns:
        tuple: (width_mm, height_mm) or None if cancelled
    """
    prompt = "Select page format (A4, A3, A2, A1) [A3]: "
    default = "A3"
    result = rs.GetString(prompt, default)

    if result is None:
        return None

    result = result.strip()
    if not result:
        result = default

    dimensions = _get_page_format_dimensions(result)
    if dimensions is None:
        logger.warning("Invalid format '%s', using default A3", result)
        dimensions = _get_page_format_dimensions("A3")
        result = "A3"

    # Prompt for orientation
    logger.info("Prompting for page orientation...")
    landscape = _prompt_orientation()
    if landscape is None:
        logger.warning(
            "Orientation selection cancelled, using default Portrait")
        landscape = False

    # Apply orientation
    if landscape:
        dimensions = (dimensions[1], dimensions[0])  # Swap width and height

    orientation_str = "Landscape" if landscape else "Portrait"
    logger.info("Selected page format: %s %s (%.1f x %.1f mm)",
                result.upper(), orientation_str, dimensions[0], dimensions[1])
    return dimensions


def _prompt_scale():
    """
    Prompt user for the drawing scale.
    Format: 1mm page = XX mm drawing

    Returns:
        tuple: (scale_paper_mm, scale_model_mm) where scale_model_mm is the drawing scale
               or None if cancelled
    """
    prompt = "Enter scale: 1mm page = XX mm drawing [200]: "
    default = "200"
    result = rs.GetString(prompt, default)

    if result is None:
        return None

    result = result.strip()
    if not result:
        result = default

    try:
        scale_model_mm = float(result)
        if scale_model_mm <= 0:
            raise ValueError("Scale must be positive")
        scale_paper_mm = 1.0
        logger.info("Scale set: 1 mm page = %.1f mm drawing", scale_model_mm)
        return (scale_paper_mm, scale_model_mm)
    except (ValueError, TypeError) as e:
        logger.warning("Invalid scale '%s', using default 200", result)
        return (1.0, 200.0)


def _prompt_folder():
    """
    Prompt user to select a folder containing DWG files.

    Returns:
        str: Folder path or None if cancelled
    """
    # Try to get folder using Rhino's file dialog
    try:
        folder = rs.BrowseForFolder(
            message="Select folder containing DWG files")
        if folder and os.path.isdir(folder):
            logger.info("Selected folder: %s", folder)
            return folder
    except Exception:
        logger.debug("BrowseForFolder failed, trying GetString", exc_info=True)

    # Fallback: prompt for folder path as string
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    prompt = "Enter folder path containing DWG files [{}]: ".format(desktop)
    result = rs.GetString(prompt, desktop)

    if result is None:
        return None

    result = result.strip()
    if not result:
        result = desktop

    # Expand user home directory if needed
    result = os.path.expanduser(result)

    if not os.path.isdir(result):
        logger.error("Invalid folder path: %s", result)
        return None

    logger.info("Selected folder: %s", result)
    return result


# ---------- helpers ----------


def _get_page_view_names():
    """
    Retrieve the names of all layout page views in the active document.

    This prefers RhinoCommon (reliable across Rhino versions). If the legacy
    rhinoscriptsyntax API `PageViewNames` exists in the current runtime, it is
    used as a soft fallback for compatibility with certain environments.

    Returns:
        list[str]: A list of layout page names (strings). May be empty.
    """
    try:
        # Primary path: RhinoCommon is stable and available in headless-safe contexts
        page_views = Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews()
        return [pv.PageName for pv in page_views] if page_views else []
    except Exception:
        logger.debug(
            "RhinoCommon page view discovery failed; trying rs.PageViewNames if present.", exc_info=True)
    # Soft fallback: only call if attribute exists to avoid AttributeError on some builds
    try:
        page_view_names_func = getattr(rs, "PageViewNames", None)
        if callable(page_view_names_func):
            names = page_view_names_func() or []
            return list(names)
    except Exception:
        logger.debug("rs.PageViewNames fallback failed.", exc_info=True)
    return []


def _activate_layout_or_create(name, width_mm=420.0, height_mm=297.0):
    """
    Activate an existing layout by name or create it if missing.

    Args:
        name: Layout page name
        width_mm: Page width in millimeters (default: 420.0)
        height_mm: Page height in millimeters (default: 297.0)

    Returns:
        tuple: (page_view, page_id) where page_view is the PageView object
               and page_id is the viewport GUID
    """
    logger.debug("Activating or creating layout: %s", name)
    # Try existing
    for pv in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
        if pv.PageName == name:
            Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = pv
            try:
                rs.CurrentView(pv.PageName)
            except Exception:
                logger.debug("Failed to switch to page view: %s",
                             pv.PageName, exc_info=True)
            return pv, pv.ActiveViewportID
    # Create new layout using RhinoCommon (more reliable than rs.AddLayout)
    try:
        # Use RhinoCommon's AddPageView which accepts name, width, height
        page_view = Rhino.RhinoDoc.ActiveDoc.Views.AddPageView(
            name, width_mm, height_mm)
        if page_view is None:
            raise Exception("Failed to create layout: {}".format(name))
        page_id = page_view.ActiveViewportID
        # Activate the newly created page view
        Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = page_view
        try:
            rs.CurrentView(page_view.PageName)
        except Exception:
            logger.debug("Failed to set CurrentView to layout %s",
                         page_view.PageName, exc_info=True)
        logger.debug("Created layout '%s' with size %.1f x %.1f mm",
                     name, width_mm, height_mm)
        return page_view, page_id
    except Exception as e:
        # Fallback: try rs.AddLayout with just the name, then set dimensions
        logger.debug(
            "RhinoCommon AddPageView failed, trying rs.AddLayout fallback", exc_info=True)
        try:
            page_id = rs.AddLayout(name)
            if not page_id:
                raise Exception("Failed to create layout: {}".format(name))
            # Resolve page view
            page_view = None
            for pv in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
                if pv.ActiveViewportID == page_id:
                    page_view = pv
                    break
            if page_view is None:
                # fallback by name match
                for pv in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
                    if pv.PageName == name:
                        page_view = pv
                        break
            if page_view is None:
                raise Exception(
                    "Created layout but could not resolve page view for: {}".format(name))
            # Set page dimensions
            try:
                page_view.PageWidth = width_mm
                page_view.PageHeight = height_mm
            except Exception:
                logger.debug(
                    "Failed to set page dimensions, using defaults", exc_info=True)
            Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = page_view
            try:
                rs.CurrentView(page_view.PageName)
            except Exception:
                logger.debug("Failed to set CurrentView to layout %s",
                             page_view.PageName, exc_info=True)
            return page_view, page_id
        except Exception:
            raise Exception("Failed to create layout '{}': {}".format(name, e))


def _duplicate_master_layout(master_name):
    """
    Duplicate master layout by name.
    Returns the new page view id (GUID) or None on failure.
    """
    logger.debug("Duplicating master layout: %s", master_name)
    try:
        return rs.DuplicateLayout(master_name)
    except Exception:
        logger.debug("Failed to duplicate master layout.", exc_info=True)
        return None


def _find_master_layout(explicit_name=None):
    """
    Decide on a master layout to use as a template.
    Priority:
      1) explicit_name (if provided and exists)
      2) layout named 'MASTER' (case-insensitive)
      3) first existing layout
      4) None (no master)
    Returns layout name or None.
    """
    # Use robust page view discovery that works across Rhino versions
    names = _get_page_view_names()
    if not names:
        logger.info(
            "No existing layouts found; proceeding without a master layout.")
        return None
    if explicit_name and explicit_name in names:
        logger.info("Using explicit master layout: %s", explicit_name)
        return explicit_name
    for n in names:
        if n.strip().lower() == "master":
            logger.info("Using 'MASTER' as master layout.")
            return n
    logger.info("Using first layout as master: %s", names[0])
    return names[0]


def _detail_rect_with_margin(page_view, margin_mm=10.0):
    """
    Compute a detail rectangle (left,bottom,right,top) within the page bounds in mm.
    """
    w = page_view.PageWidth
    h = page_view.PageHeight
    l = margin_mm
    b = margin_mm
    r = max(margin_mm, w - margin_mm)
    t = max(margin_mm, h - margin_mm)
    return l, b, r, t


def _resolve_layout_guid(page_view):
    """
    Best-effort resolution of the layout (page view) GUID required by rs.AddDetail.
    Tries multiple strategies to obtain a GUID for the provided page_view.
    Returns GUID or None if not found.
    """
    # Try rhinoscriptsyntax ViewId by page name (works for page views in many builds)
    try:
        vid = rs.ViewId(page_view.PageName)
        if vid:
            return vid
    except Exception:
        logger.debug("rs.ViewId failed for page '%s'", getattr(
            page_view, "PageName", "?"), exc_info=True)
    # Try RhinoCommon: some builds expose MainViewport.Id which is a Guid we can reuse
    try:
        for pv in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
            if pv.PageName == page_view.PageName:
                try:
                    mv = pv.MainViewport
                    if mv and getattr(mv, "Id", None):
                        return mv.Id
                except Exception:
                    pass
    except Exception:
        logger.debug(
            "Failed to resolve layout GUID via RhinoCommon.", exc_info=True)
    return None


def _add_or_replace_single_detail(page_view, margin_mm=10.0):
    """
    Ensure the layout has exactly one unlocked detail covering the page with a margin.
    Returns the detail object id (GUID).
    """
    logger.debug("Preparing single detail on page: %s", page_view.PageName)
    # Delete existing details
    try:
        detail_ids = rs.ObjectsByType(32768, select=False) or [
        ]  # 32768 = detail objects
        # Filter by current page
        current_page_id = page_view.ActiveViewportID
        to_delete = []
        for did in detail_ids:
            try:
                obj = Rhino.RhinoDoc.ActiveDoc.Objects.Find(did)
                if obj and obj.Attributes and obj.Attributes.Space == Rhino.DocObjects.ActiveSpace.PageSpace:
                    if obj.Attributes.LayoutIndex == page_view.PageNumber:
                        to_delete.append(did)
            except Exception:
                logger.debug(
                    "Error inspecting detail object; skipping.", exc_info=True)
        if to_delete:
            rs.DeleteObjects(to_delete)
    except Exception:
        logger.debug("Failed to clean prior details.", exc_info=True)
    # Add new detail
    l, b, r, t = _detail_rect_with_margin(page_view, margin_mm=margin_mm)
    logger.debug(
        "Detail rectangle: left=%.1f, bottom=%.1f, right=%.1f, top=%.1f", l, b, r, t)

    # Ensure we're on the correct page view
    try:
        Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = page_view
        rs.CurrentView(page_view.PageName)
        time.sleep(0.1)  # Give Rhino time to switch views
    except Exception as e:
        logger.debug(
            "Failed to activate page view before adding detail: %s", e, exc_info=True)

    # Try rs.AddDetail first (most reliable)
    detail_id = None
    try:
        # Ensure page view is active before adding detail
        Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = page_view
        rs.CurrentView(page_view.PageName)
        time.sleep(0.15)  # Give Rhino time to fully activate the view

        # rs.AddDetail expects the layout GUID. Resolve best-effort.
        layout_guid = _resolve_layout_guid(page_view)
        if not layout_guid:
            raise Exception(
                "Could not resolve layout GUID for page '{}'".format(page_view.PageName))
        logger.debug("Calling rs.AddDetail with: layout_id=%s, corner1=(%.1f, %.1f), corner2=(%.1f, %.1f)",
                     str(layout_guid), l, b, r, t)
        # rs.AddDetail expects two 2D points (corner1, corner2)
        corner1 = (float(l), float(b))
        corner2 = (float(r), float(t))
        # Use positional parameters only for compatibility across builds
        detail_id = rs.AddDetail(layout_guid, corner1, corner2)
        if detail_id:
            logger.debug("Created detail using rs.AddDetail: %s", detail_id)
            # Unlock to allow zoom and scale
            try:
                rs.DetailLock(detail_id, False)
            except Exception:
                pass
            # Try to set the detail view to Top (World) for consistent orientation
            try:
                try:
                    rs.CurrentDetail(detail_id, True)
                except Exception:
                    logger.debug(
                        "rs.CurrentDetail activation failed.", exc_info=True)
                rs.Command(u'_-SetView _World _Top _Enter', echo=False)
            except Exception:
                logger.debug(
                    "Failed to set detail view to World Top.", exc_info=True)
            return detail_id
        else:
            logger.warning(
                "rs.AddDetail returned None for page: %s (this usually means the call failed silently)", page_view.PageName)
    except Exception as e:
        logger.error("rs.AddDetail failed with exception for page '%s': %s",
                     page_view.PageName, e, exc_info=True)

    # Fallback: Try using command-based approach (less reliable, but worth trying)
    if not detail_id:
        try:
            logger.debug("Trying command-based detail creation as last resort")
            # Note: Command-based detail creation is unreliable and may not work
            # We'll try it but don't expect it to succeed in all cases
            # The Detail command typically requires interactive input for coordinates
            pass  # Skip command-based approach as it's causing parsing errors
        except Exception as e:
            logger.debug(
                "Command-based detail creation skipped: %s", e)

    # Last resort: Try RhinoCommon
    if not detail_id:
        try:
            logger.debug("Trying RhinoCommon AddDetail")
            # Create detail using RhinoCommon - need to convert to proper coordinate system
            # Details on page space use page coordinates (mm from bottom-left)
            plane = Rhino.Geometry.Plane(
                Rhino.Geometry.Point3d(l, b, 0),
                Rhino.Geometry.Vector3d.XAxis,
                Rhino.Geometry.Vector3d.YAxis
            )
            detail = Rhino.RhinoDoc.ActiveDoc.Views.AddDetail(
                page_view.Id,
                Rhino.Geometry.Rectangle3d(
                    plane,
                    Rhino.Geometry.Interval(0, r - l),
                    Rhino.Geometry.Interval(0, t - b)
                ),
                "Top"
            )
            if detail:
                detail_id = detail.Id
                logger.debug("Created detail using RhinoCommon: %s", detail_id)
                try:
                    rs.DetailLock(detail_id, False)
                except Exception:
                    pass
                return detail_id
        except Exception as e:
            logger.debug("RhinoCommon AddDetail failed: %s", e, exc_info=True)

    logger.error("All methods failed to add detail to page: %s",
                 page_view.PageName)
    return None


def _set_detail_scale(detail_id, paper_mm_per_model_unit=1.0/0.2, model_unit="Meters"):
    """
    Configure detail scale using rs.DetailScale:
      paper_length(mm) : model_length(model_unit)

    Args:
        detail_id: GUID of the detail viewport
        paper_mm_per_model_unit: Ratio of paper_mm / model_length_in_model_unit
                                 Example: if 1mm paper = 0.2m model, then this is 1.0/0.2 = 5.0
        model_unit: Unit string for the model (default: "Meters")

    The scale is set as: paper_length(mm) : model_length(model_unit)
    """
    try:
        # Calculate model_length from the ratio
        # paper_mm_per_model_unit = paper_mm / model_length
        # So: model_length = paper_mm / paper_mm_per_model_unit
        # We use paper_length = 1.0 mm as the reference
        paper_length = 1.0
        model_length = paper_length / paper_mm_per_model_unit

        # Explicitly set the scale with units; rs.DetailScale expects model_length, paper_length, and optional units
        rs.DetailScale(detail_id, model_length, paper_length,
                       model_unit, "Millimeters")
        logger.debug("Detail scale set: %.3f mm paper = %.3f %s model",
                     paper_length, model_length, model_unit)
        return True
    except Exception:
        logger.debug("Failed to set detail scale.", exc_info=True)
        return False


def _import_dwg_capture_new_objects(path):
    """
    Import a DWG and return the list of newly added object ids.
    """
    logger.info("Importing DWG: %s", path)
    before = set(rs.AllObjects() or [])
    # Use robust command variants to avoid UI prompts
    cmds = [
        u'-_Import "{}" _Enter'.format(path),
        u'-_Import {} _Enter'.format(path),
    ]
    for cmd in cmds:
        rs.Command(cmd, echo=False)
        time.sleep(0.15)
        after = set(rs.AllObjects() or [])
        diff = list(after - before)
        if diff:
            logger.info("Imported %d new objects from %s",
                        len(diff), os.path.basename(path))
            return diff
    logger.error("DWG import produced no new geometry: %s", path)
    return []


def _zoom_selected_in_detail(detail_id, object_ids):
    """
    Activate the detail, select the given objects, perform a Zoom Selected, then unselect.
    """
    if not object_ids:
        return
    try:
        # Activate layout view first
        parent_layout = rs.coercerhinoobject(detail_id).Attributes.LayoutIndex
    except Exception:
        parent_layout = None
    try:
        rs.SelectObjects(object_ids)
    except Exception:
        pass
    # Attempt to activate the detail and zoom to selection
    try:
        # Use API to activate the detail if available
        try:
            rs.CurrentDetail(detail_id, True)
        except Exception:
            # Fallback: try without the activate flag
            try:
                rs.CurrentDetail(detail_id)
            except Exception:
                pass
        rs.Command('_-Zoom _Selected _Enter', echo=False)
    except Exception:
        logger.debug("Failed to zoom selected within detail.", exc_info=True)
    try:
        rs.UnselectAllObjects()
    except Exception:
        pass


def _export_pdf_all_layouts(out_pdf_path):
    """
    Attempt a vector PDF export of all layouts to the specified file path.
    Tries multiple command variants for macOS builds.
    Returns True on success.
    """
    logger.info(
        "Attempting vector PDF export of all layouts -> %s", out_pdf_path)
    # Remove stale file
    try:
        if os.path.exists(out_pdf_path):
            os.remove(out_pdf_path)
    except Exception:
        logger.debug("Failed to remove stale PDF file (continuing): %s",
                     out_pdf_path, exc_info=True)
    cmds = [
        # RhinoPDF (preferred)
        u'-_Print _Setup _Destination=_PDF _OutputColor=_Display _OutputType=_Vector _Enter _View=_AllLayouts _Enter _Enter',
        # Sometimes requiring extra _Enter on mac
        u'-_Print _Destination=_PDF _OutputType=_Vector _Enter _View=_AllLayouts _Enter _Enter',
    ]
    for cmd in cmds:
        rs.Command(cmd, echo=False)
        time.sleep(0.5)
        # Try to force output path (some builds honor this, some ignore)
        try:
            rs.Command(
                u'-_Print _OutputFile "{}" _Enter _Go'.format(out_pdf_path), echo=False)
        except Exception:
            pass
        time.sleep(0.5)
        try:
            if os.path.exists(out_pdf_path) and os.path.getsize(out_pdf_path) > 0:
                logger.info("PDF export succeeded: %s", out_pdf_path)
                return True
        except Exception:
            logger.debug("PDF file check failed.", exc_info=True)
    logger.warning("Bulk PDF export failed; will try per-layout export.")
    return False


def _export_pdf_per_layout(output_dir, basename_prefix="Sheet"):
    """
    Export each layout as an individual PDF to output_dir.
    Returns list of generated file paths.
    """
    logger.info("Exporting per-layout PDFs to: %s", output_dir)
    paths = []
    page_views = Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews()
    idx = 1
    for pv in page_views:
        try:
            Rhino.RhinoDoc.ActiveDoc.Views.ActiveView = pv
            rs.CurrentView(pv.PageName)
        except Exception:
            logger.debug("Failed to activate page before export: %s",
                         pv.PageName, exc_info=True)
        outpath = os.path.join(
            output_dir, "{}-{:02d}-{}.pdf".format(basename_prefix, idx, pv.PageName))
        # Try export using Print command with PDF destination
        try:
            if os.path.exists(outpath):
                os.remove(outpath)
        except Exception:
            pass

        # Use Print command with PDF output for individual layouts
        # Try multiple command syntaxes for compatibility
        export_success = False

        # Single command with proper quoting (step-by-step variants are not supported)
        try:
            logger.debug(
                "Trying single-line Print command for view: %s", pv.PageName)
            safe_path = outpath.replace('\\', '\\\\').replace('"', '\\"')
            safe_view = pv.PageName.replace('"', '\\"')
            cmd = u'-_Print _Destination=_PDF _OutputType=_Vector _View="{}" _OutputFile="{}" _Go'.format(
                safe_view, safe_path)
            logger.debug("Executing: %s", cmd)
            rs.Command(cmd, echo=False)
            time.sleep(1.0)
            if os.path.exists(outpath) and os.path.getsize(outpath) > 0:
                logger.info("Exported layout PDF: %s", outpath)
                paths.append(outpath)
                export_success = True
        except Exception as e:
            logger.debug("Single-line Print command failed: %s",
                         e, exc_info=True)

        # Variant 3: Try using RhinoCommon Print functionality if available
        if not export_success:
            try:
                logger.debug("Trying RhinoCommon print functionality")
                # Note: Direct PDF export via RhinoCommon may require additional setup
                # For now, we rely on the Print command variants above
                pass
            except Exception as e:
                logger.debug("RhinoCommon print approach not available: %s", e)

        if not export_success:
            logger.warning(
                "Failed to export PDF for layout: %s (path: %s)", pv.PageName, outpath)
        idx += 1
    return paths

# ---------- main orchestration ----------


def assemble_from_dwgs(
    dwg_paths=None,
    dwg_folder=None,
    scale_paper_mm=1.0,
    scale_model_mm=200.0,
    master_layout_name=None,
    page_width_mm=420.0,
    page_height_mm=297.0,
    margin_mm=10.0,
    output_pdf_path=None
):
    """
    Assemble one layout per DWG:
    - dwg_paths: list of file paths. If None, discovers DWG files from dwg_folder
    - dwg_folder: folder path to search for DWG files. If None and dwg_paths is None,
                  defaults to '~/Desktop' searching for '*-Export.dwg'
    - scale_paper_mm: paper length (mm) in the scale definition (default 1.0)
    - scale_model_mm: model/drawing length (mm) corresponding to scale_paper_mm (default 200.0)
                      This represents "1mm page = XX mm drawing"
    - master_layout_name: optional explicit master layout to duplicate
    - page_width_mm, page_height_mm: page size for created layouts
    - margin_mm: detail margin from edges
    - output_pdf_path: if provided, attempts to export combined vector PDF here;
      otherwise exports individual PDFs to the Desktop.
    Returns a dict with keys: 'layouts', 'pdf' or 'pdfs'
    """
    # Discover DWG files if not provided
    if dwg_paths is None:
        if dwg_folder is None:
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            dwg_folder = desktop
        if not os.path.isdir(dwg_folder):
            raise Exception("Invalid folder path: {}".format(dwg_folder))
        # Search for all .dwg files in the folder
        dwg_paths = sorted(glob.glob(os.path.join(dwg_folder, "*.dwg")))
        if not dwg_paths:
            # Fallback to old pattern if no DWGs found
            dwg_paths = sorted(
                glob.glob(os.path.join(dwg_folder, "*-Export.dwg")))

    if not dwg_paths:
        raise Exception("No DWG files provided or discovered in folder: {}".format(
            dwg_folder or "specified"))
    logger.info("Found %d DWGs to assemble.", len(dwg_paths))

    # Determine master layout (if any)
    master = _find_master_layout(explicit_name=master_layout_name)

    created_layouts = []
    for path in dwg_paths:
        name_base = os.path.splitext(os.path.basename(path))[0]
        # Build page: either duplicate master or create fresh
        if master:
            dup_id = _duplicate_master_layout(master)
            if dup_id:
                # Rename the duplicated page to match DWG base
                try:
                    # dup_id is the page (layout) GUID; pass directly
                    rs.RenameLayout(dup_id, name_base)
                except Exception:
                    logger.debug(
                        "Failed to rename duplicated layout; will set active and proceed.", exc_info=True)
                # Resolve page view by duplicated layout's viewport id first
                pv = None
                try:
                    for v in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
                        # Match by ActiveViewportID if available
                        try:
                            if getattr(v, "ActiveViewportID", None) == dup_id:
                                pv = v
                                break
                        except Exception:
                            pass
                    # Fallback to name match after rename
                    if pv is None:
                        for v in Rhino.RhinoDoc.ActiveDoc.Views.GetPageViews():
                            if v.PageName == name_base:
                                pv = v
                                break
                except Exception:
                    logger.debug(
                        "Error resolving page view after duplication.", exc_info=True)
                if pv is None:
                    # fallback: try to find by active viewport ID or create new
                    try:
                        active_view = Rhino.RhinoDoc.ActiveDoc.Views.ActiveView
                        # Only use if it's actually a page view
                        if isinstance(active_view, Rhino.Display.RhinoPageView):
                            pv = active_view
                        else:
                            # Create new layout if we can't find the duplicated one
                            logger.warning(
                                "Could not resolve duplicated layout, creating new one")
                            pv, _ = _activate_layout_or_create(
                                name_base, width_mm=page_width_mm, height_mm=page_height_mm)
                    except Exception:
                        # Create new layout as last resort
                        logger.warning(
                            "Failed to resolve duplicated layout, creating new one")
                        pv, _ = _activate_layout_or_create(
                            name_base, width_mm=page_width_mm, height_mm=page_height_mm)
                page_view = pv
            else:
                page_view, _ = _activate_layout_or_create(
                    name_base, width_mm=page_width_mm, height_mm=page_height_mm)
        else:
            page_view, _ = _activate_layout_or_create(
                name_base, width_mm=page_width_mm, height_mm=page_height_mm)

        # Ensure a single detail
        # Validate that we have a valid page view
        if not isinstance(page_view, Rhino.Display.RhinoPageView):
            logger.error(
                "Invalid page view type for layout creation. Expected RhinoPageView, got: %s", type(page_view))
            # Try to recreate the layout
            try:
                page_view, _ = _activate_layout_or_create(
                    name_base, width_mm=page_width_mm, height_mm=page_height_mm)
            except Exception as e:
                logger.error("Failed to recreate layout %s: %s", name_base, e)
                continue

        detail_id = _add_or_replace_single_detail(
            page_view, margin_mm=margin_mm)
        if not detail_id:
            logger.error("Could not create a detail on layout: %s",
                         page_view.PageName)
            continue

        # Import DWG and capture new objects
        new_ids = _import_dwg_capture_new_objects(path)
        if not new_ids:
            logger.error("Skipping layout due to empty import: %s", path)
            continue

        # Set requested scale
        # Convert scale_model_mm (mm) to meters for the detail scale function
        # scale_model_m = scale_model_mm / 1000.0
        try:
            # _set_detail_scale expects: paper_mm_per_model_unit = scale_paper_mm / scale_model_m
            # where scale_model_m is in meters
            # User input: 1mm page = scale_model_mm mm drawing
            # So: scale_model_m = scale_model_mm / 1000.0 (convert mm to meters)
            scale_model_m = scale_model_mm / 1000.0
            _set_detail_scale(
                detail_id, paper_mm_per_model_unit=scale_paper_mm/scale_model_m, model_unit="Meters")
        except Exception:
            logger.debug(
                "Detail scale helper failed, continuing.", exc_info=True)

        # Zoom selected objects inside the detail
        _zoom_selected_in_detail(detail_id, new_ids)

        # Lock the detail to prevent accidental changes
        try:
            rs.DetailLock(detail_id, True)
        except Exception:
            pass

        created_layouts.append(page_view.PageName)
        logger.info("Prepared layout: %s", page_view.PageName)

    # Export PDF(s)
    desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
    if output_pdf_path is None:
        output_pdf_path = os.path.join(desktop, "AssembledLayouts.pdf")

    pdf_ok = _export_pdf_all_layouts(output_pdf_path)
    result = {"layouts": created_layouts}
    if pdf_ok:
        result["pdf"] = output_pdf_path
    else:
        # Fallback to per-layout exports
        pdfs = _export_pdf_per_layout(desktop, basename_prefix="Assembled")
        result["pdfs"] = pdfs
    return result


def main():
    """
    Entry point for interactive runs inside Rhino.
    Prompts user for:
      - Page format (A4, A3, A2, A1)
      - Drawing scale (1mm page = XX mm drawing)
      - Folder containing DWG files
    Exports combined PDF to Desktop/AssembledLayouts.pdf (or per-layout on fallback)
    """
    try:
        rs.EnableRedraw(False)
    except Exception:
        pass

    try:
        # Prompt for page format
        logger.info("Prompting for page format...")
        page_dims = _prompt_page_format()
        if page_dims is None:
            logger.warning("Page format selection cancelled, using default A3")
            page_dims = _get_page_format_dimensions("A3")
        page_width_mm, page_height_mm = page_dims

        # Prompt for scale
        logger.info("Prompting for drawing scale...")
        scale_params = _prompt_scale()
        if scale_params is None:
            logger.warning("Scale input cancelled, using default 200")
            scale_params = (1.0, 200.0)
        scale_paper_mm, scale_model_mm = scale_params

        # Prompt for folder
        logger.info("Prompting for DWG folder...")
        dwg_folder = _prompt_folder()
        if dwg_folder is None:
            logger.error("Folder selection cancelled. Aborting.")
            return

        # Run assembly with user-provided parameters
        result = assemble_from_dwgs(
            dwg_paths=None,
            dwg_folder=dwg_folder,
            scale_paper_mm=scale_paper_mm,
            scale_model_mm=scale_model_mm,
            master_layout_name=None,
            page_width_mm=page_width_mm,
            page_height_mm=page_height_mm,
            margin_mm=10.0,
            output_pdf_path=None
        )
        logger.info("Assembly complete: %s", result)
    except Exception as e:
        logger.exception("Assembly failed: %s", e)
    finally:
        try:
            rs.EnableRedraw(True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
