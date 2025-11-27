"""
Export Letterhead Utilities

This module provides helper functions for adding letterhead to Excel/CSV exports.
It handles settings retrieval, template rendering, and letterhead row generation.

Key Functions:
- _get_settings(): Retrieves and validates export letterhead settings
- _build_context(): Builds context dictionary for template rendering with available variables
- _render_template(): Renders Jinja2 templates with context
- _generate_letterhead_rows(): Generates letterhead rows from template

Template Variables Available:
- company: Company name from user defaults
- doctype: Document type identifier (list view exports) or fallback name
- report_name: Report title (e.g., "General Ledger") when available
- user_fullname: Current user's full name
- date: Current date (datetime.date)
- time: Current time (datetime.time)
- now: Current datetime object (datetime.datetime)
- frappe: Frappe object for advanced scripting (database lookups, etc.)

Template Examples:
1. Simple header:
   {{ company }}
   Export Report

2. Multi-column with separator (| or tab):
   {{ company }} | {{ report_name or doctype }} | {{ date }}

3. Conditional logic:
   {% if report_name %}
   Report: {{ report_name }}
   {% else %}
   Export: {{ doctype }}
   {% endif %}
"""

import frappe
from frappe.utils import now_datetime


def _safe_get_value(source, *keys):
    """
    Safely fetch a value from dict-like or object-like sources.

    Handles frappe._dict instances (which return None instead of raising AttributeError)
    and plain dicts/objects. Returns the first non-empty value found for the provided keys.
    """
    if not source:
        return None

    for key in keys:
        value = None
        if isinstance(source, dict):
            value = source.get(key)
        else:
            try:
                value = getattr(source, key)
            except AttributeError:
                value = None

        if isinstance(value, str):
            value = value.strip()

        if value not in (None, ""):
            return value

    return None


def _get_settings():
    """
    Retrieve and validate Export Letterhead Settings.
    
    Returns a dictionary with the following keys:
    - enabled: Boolean - Whether letterhead is enabled
    - letterhead_template: String - Jinja2 template for letterhead
    - font_name: String - Font name to use (e.g., "Arial", "Calibri")
    - font_size: Integer - Font size (1-409)
    - add_printed_by: Boolean - Whether to add "Printed by" row
    
    Returns None if settings cannot be retrieved.
    """
    try:
        s = frappe.get_single("Export Letterhead Settings")
        
        # Get font_size and ensure it's an integer
        font_size = getattr(s, "font_size", 11)
        try:
            font_size = int(font_size) if font_size else 11
        except (ValueError, TypeError):
            font_size = 11
        
        # Get font_name and ensure it's a valid string
        font_name = getattr(s, "font_name", "Arial")
        if font_name:
            font_name = str(font_name).strip()
        if not font_name:
            font_name = "Arial"
        
        # Clean font name (remove invalid characters)
        import re
        font_name = re.sub(r'[^\w\s\-]', '', font_name).strip()
        if not font_name:
            font_name = "Arial"
        
        return {
            "enabled": getattr(s, "enabled", False),
            "letterhead_template": getattr(s, "letterhead_template", ""),
            "font_name": font_name,
            "font_size": font_size,
            "add_printed_by": getattr(s, "add_printed_by", True),
        }
    except Exception as e:
        frappe.logger("letterhead").debug("Failed to get export letterhead settings", exc_info=True)
        return None


def _build_context(args, kwargs):
    """
    Build context dictionary for Jinja2 template rendering.
    
    Extracts available information from function arguments and Frappe session
    to provide variables for template rendering.
    
    Args:
        args: Tuple of positional arguments (unused but kept for compatibility)
        kwargs: Dictionary of keyword arguments (may contain doctype, report_name)
    
    Returns:
        Dictionary with template variables:
        - frappe: Frappe object
        - user_fullname: Current user's full name
        - doctype: Document type identifier or filename
        - report_name: Report title (if available)
        - company: Company name (user default)
        - date: Current date
        - time: Current time
        - now: Current datetime
    """
    context = {
        "frappe": frappe,
    }
    
    try:
        context["user_fullname"] = frappe.session.user_fullname or frappe.session.user
    except Exception:
        context["user_fullname"] = frappe.session.user if hasattr(frappe.session, 'user') else None

    if isinstance(kwargs, dict):
        if kwargs.get("doctype"):
            context["doctype"] = kwargs.get("doctype")
        if kwargs.get("report_name"):
            context["report_name"] = kwargs.get("report_name")

    # Fallback to request form_dict (helps for older Frappe versions / background jobs)
    form_dict = getattr(frappe.local, "form_dict", None)
    if not context.get("doctype"):
        context["doctype"] = _safe_get_value(form_dict, "doctype", "ref_doctype", "data_doctype")
    if not context.get("report_name"):
        context["report_name"] = _safe_get_value(form_dict, "report_name", "report", "title")
    
    # Prefer report_name, fall back to doctype if missing
    if context.get("report_name") and not context.get("doctype"):
        context["doctype"] = context["report_name"]
    elif context.get("doctype") and not context.get("report_name"):
        context["report_name"] = context["doctype"]
    
    # Company from user defaults
    try:
        context["company"] = frappe.defaults.get_user_default("company") or ""
    except Exception:
        context["company"] = ""
    
    # Add common variables
    try:
        from frappe.utils import now_datetime
        now = now_datetime()
        context["now"] = now
        context["date"] = now.date()
        context["time"] = now.time()
    except Exception:
        import datetime
        now = datetime.datetime.now()
        context["now"] = now
        context["date"] = now.date()
        context["time"] = now.time()
    
    return context


def _render_template(template_text, context):
    """
    Render Jinja2 template with provided context.
    
    Uses Jinja2 for template rendering. Falls back to frappe.render_template
    or returns original text if rendering fails.
    
    Args:
        template_text: Jinja2 template string
        context: Dictionary of variables for template rendering
    
    Returns:
        Rendered template string
    
    Example:
        template = "{{ company }} | {{ date }}"
        context = {"company": "Acme Corp", "date": "2025-01-15"}
        Returns: "Acme Corp | 2025-01-15"
    """
    try:
        from jinja2 import Template
        tpl = Template(template_text)
        return tpl.render(**(context or {}))
    except Exception:
        # Fallback to frappe.render_template if available
        try:
            return frappe.render_template(template_text, context)
        except Exception:
            return template_text


def _generate_letterhead_rows(settings, context=None):
    """
    Generate letterhead rows from template for Excel/CSV exports.
    
    Processes the letterhead template, renders it with context variables,
    and converts it into rows of cells. Each line in the template becomes
    a row. Columns can be separated by tabs (\t) or pipes (|).
    
    Args:
        settings: Dictionary with settings (enabled, letterhead_template, add_printed_by)
        context: Optional context dictionary for template variables
    
    Returns:
        List of rows, where each row is a list of cell values.
        Empty list if disabled or no template.
    
    Template Format:
        - Each line = one row
        - Use | or tab to separate columns
        - Empty lines are ignored
    
    Examples:
        1. Single column rows:
           Template: "{{ company }}\nExport Report"
           Returns: [["Acme Corp"], ["Export Report"]]
        
        2. Multi-column with pipe:
           Template: "{{ company }} | {{ date }}"
           Returns: [["Acme Corp", "2025-01-15"]]
        
        3. Multi-column with tab:
           Template: "{{ company }}\t{{ user_fullname }}"
           Returns: [["Acme Corp", "John Doe"]]
        
        4. With "Printed by" row (if enabled):
           Returns: [
               ["Acme Corp"],
               ["Printed by: John Doe", "Date: 2025-01-15", "Time: 14:30:00"]
           ]
    """
    # Check if letterhead is enabled and template exists
    if not settings or not settings.get("enabled") or not settings.get("letterhead_template"):
        return []
    
    template_text = settings.get("letterhead_template", "")
    if not template_text:
        return []
    
    # Build context with additional info if not provided
    if context is None:
        context = {}
    
    # Ensure report_name falls back to doctype if not already set
    if not context.get("report_name") and context.get("doctype"):
        context["report_name"] = context["doctype"]
    
    # Add common context variables (these are always available)
    context.update({
        "frappe": frappe,
        "user_fullname": frappe.session.user_fullname or frappe.session.user,
        "now": now_datetime(),
        "date": now_datetime().date(),
        "time": now_datetime().time(),
        "company": frappe.defaults.get_user_default("company") or "",
    })
    
    # Render template with context
    rendered = _render_template(template_text, context)
    
    # Split rendered template into rows (each line becomes a row)
    rows = []
    for line in rendered.split('\n'):
        if line.strip():  # Skip empty lines
            # Each row can have multiple columns separated by tabs or pipes
            if '\t' in line:
                # Tab-separated columns
                cells = line.split('\t')
            elif '|' in line:
                # Pipe-separated columns (strip whitespace from each cell)
                cells = [c.strip() for c in line.split('|')]
            else:
                # Single cell row (no separator found)
                cells = [line]
            rows.append(cells)
    
    # Add "Printed by" row if enabled (shows who exported and when)
    if settings.get("add_printed_by", True):
        printed_by_row = [
            f"Printed by: {context.get('user_fullname', frappe.session.user)}",
            f"Date: {context.get('date', '')}",
            f"Time: {context.get('time', '')}"
        ]
        rows.append(printed_by_row)
    
    # Always add a blank row at the end to visually separate header from data
    if rows:
        rows.append([""])
    
    return rows
