# ============================================================
#  cid_compare_core.py  —  Custom Comparison Engine Template
#
#  This file is OPTIONAL. If present in the same folder as
#  cid_gui.py, it will be used instead of the built-in engine.
#
#  Use this template to plug in your own comparison logic.
#  The GUI expects the function signature below.
# ============================================================


def compare_cid_files(old_path, new_path, tech_prefix,
                      old_ied=None, new_ied=None):
    """
    Compare two IEC 61850 CID or SCD files and return differences.

    Parameters
    ----------
    old_path    : str  — full path to the OLD CID/SCD file
    new_path    : str  — full path to the NEW CID/SCD file
    tech_prefix : str  — technical key prefix (e.g. 'Q01A'),
                         auto-derived from IED name by the GUI
    old_ied     : str or None — IED name selected for the old file
    new_ied     : str or None — IED name selected for the new file

    Returns
    -------
    list of dict, each dict having:
        {
            'category': str,   # e.g. 'DataSet', 'ReportControl'
            'element':  str,   # human-readable element description
            'old':      str,   # old value  (use 'NOT PRESENT' for additions)
            'new':      str,   # new value  (use 'REMOVED' for deletions)
        }

    Categories used by the GUI for colour-coding:
        'Config Revision'
        'IP / Network'
        'DataSet'
        'Signal / FCDA'
        'ReportControl'
        'Logical Node'
    """

    differences = []

    # ── Add your comparison logic here ───────────────────────
    #
    # Example:
    #
    # import xml.etree.ElementTree as ET
    # old_tree = ET.parse(old_path)
    # new_tree = ET.parse(new_path)
    # ...
    # differences.append({
    #     'category': 'DataSet',
    #     'element':  "DS 'DS_PROT_01' → PTOC/Op fc=ST",
    #     'old':      'NOT PRESENT',
    #     'new':      'EXISTS'
    # })
    #
    # ─────────────────────────────────────────────────────────

    return differences
