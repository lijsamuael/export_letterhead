# Export Letterhead

Add customizable letterhead headers and font formatting to all Excel and CSV exports in Frappe/ERPNext.

## Features

- ✅ **Custom Letterhead**: Add 2-3 rows of letterhead to every export
- ✅ **Template Support**: Use Jinja2 templates with dynamic variables
- ✅ **Font Control**: Control font name and size for entire export
- ✅ **Auto Timestamp**: Automatically add "Printed by" row with user, date, and time
- ✅ **Smart Spacing**: Automatically inserts a blank separator row before export data
- ✅ **Works Everywhere**: Applies to all Excel/CSV exports (reports, list views, data exports)

## Installation

### Step 1: Get the App

```bash
cd $PATH_TO_YOUR_BENCH
bench get-app https://github.com/lijsamuael/export_letterhead/
```

### Step 2: Install the App

```bash
bench --site {your_site_name} install-app export_letterhead
```

### Step 3: Restart Bench (if needed)

```bash
bench restart
```

## Configuration Guide

### Step 1: Access Settings

1. Log in to your Frappe/ERPNext instance
2. Go to **Search** (press `/` or click search icon)
3. Type: `Export Letterhead Settings`
4. Click on **Export Letterhead Settings**

### Step 2: Enable the Feature

1. Check the box **"Enable Export Letterhead"**
2. This activates the letterhead functionality

### Step 3: Configure Letterhead Template

In the **"Letterhead Template"** field, enter your template using Jinja2 syntax.

**Template Rules:**
- Each line = one row in the export
- Use `|` (pipe) or `Tab` to separate columns
- Empty lines are ignored
- The app automatically inserts a blank separator row after the letterhead

**Available Variables:**
- `{{ company }}` - Company name
- `{{ doctype }}` - Document type identifier or filename
- `{{ report_name }}` - Report title (e.g., `General Ledger`) when available
- `{{ user_fullname }}` - Current user's full name
- `{{ date }}` - Current date (datetime.date)
- `{{ time }}` - Current time (datetime.time)
- `{{ now }}` - Current datetime (datetime.datetime)
- `{{ frappe }}` - Frappe object for advanced scripting
- Date/time objects support `strftime`, e.g., `{{ date.strftime('%B %d, %Y') }}` or `{{ time.strftime('%I:%M %p') }}`

**Template Examples:**

**Example 1: Simple Header**
```
{{ company }}
Export Report
```

**Example 2: Multi-Column with Pipe**
```
{{ company }} | {{ report_name or doctype }} | {{ date.strftime('%B %d, %Y') }}
```

**Example 3: Multi-Column with Tab**
```
{{ company }}	{{ user_fullname }}	{{ time.strftime('%I:%M %p') }}
```

**Example 4: With Conditional Logic**
```
{% if report_name %}
Report: {{ report_name }}
{% else %}
Export: {{ doctype }}
{% endif %}
```

**Example 5: Complex Template**
```
{{ company }}
{{ report_name or doctype }} Export
Generated on: {{ date.strftime('%B %d, %Y') }} at {{ time.strftime('%I:%M %p') }}
Exported by: {{ user_fullname }}
```

### Step 4: Configure Font Settings

**Font Name:**
- Select from dropdown: Arial, Helvetica, Times New Roman, Courier New, Verdana, Calibri, Georgia, Comic Sans MS, Impact, Lucida Console
- Default: Arial

**Font Size:**
- Enter a number (1-409)
- Default: 11
- This applies to **ALL rows** in the export (letterhead + data)

### Step 5: Configure "Printed by" Row

- **"Add 'Printed by' Row"**: Check/uncheck to enable/disable
- When enabled, automatically adds a row showing:
  - Printed by: [User Name]
  - Date: [Current Date]
  - Time: [Current Time]

### Step 6: Save Settings

Click **Save** to apply your configuration.

## Usage

Once configured, the letterhead will automatically appear on **all** Excel and CSV exports:

1. **Query Reports**: Export any query report → Letterhead added
2. **List Views**: Export any list view → Letterhead added
3. **Data Exports**: Export any doctype data → Letterhead added
4. **Custom Reports**: Any custom report export → Letterhead added

**No additional steps required!** Just export as you normally would.

## Step-by-Step Example

Let's create a letterhead for Sales Invoice exports:

### Step 1: Open Settings
- Search: `Export Letterhead Settings`
- Open the document

### Step 2: Enable
- ✅ Check "Enable Export Letterhead"

### Step 3: Create Template
Enter this in "Letterhead Template":
```
{{ company }}
Sales Invoice Export
Generated: {{ date }} at {{ time }}
```

### Step 4: Set Font
- Font Name: `Calibri`
- Font Size: `12`

### Step 5: Enable Printed By
- ✅ Check "Add 'Printed by' Row"

### Step 6: Save
- Click **Save**

### Step 7: Test
1. Go to any Sales Invoice report
2. Click Export → Excel
3. Open the file
4. You should see:
   - Your company name
   - "Sales Invoice Export"
   - Generated date/time
   - "Printed by: [Your Name]" with date and time
   - All rows using Calibri 12pt font

## Template Tips

### Working with Report Names

```jinja2
{% if report_name %}
Report: {{ report_name }}
{% else %}
Export: {{ doctype }}
{% endif %}
```

### Using Company Information

```jinja2
{{ company }}
Company Address: {{ frappe.db.get_value("Company", company, "address_line1") }}
```

### Conditional Formatting

```jinja2
{% if report_name == "Sales Invoice" %}
Sales Invoice Report
{% elif report_name == "Purchase Invoice" %}
Purchase Invoice Report
{% else %}
{{ report_name or doctype }} Export
{% endif %}
```

### Multi-Column Layout

Use pipe (`|`) for aligned columns:
```jinja2
{{ company }} | Phone: {{ frappe.db.get_value("Company", company, "phone_no") }} | Email: {{ frappe.db.get_value("Company", company, "email") }}
```

## Troubleshooting

### Letterhead Not Appearing

1. **Check if enabled**: Verify "Enable Export Letterhead" is checked
2. **Check template**: Ensure template field is not empty
3. **Clear cache**: `bench clear-cache`
4. **Restart**: `bench restart`

### Font Not Applying

1. **Check font name**: Ensure font name is valid (use dropdown options)
2. **Check font size**: Must be between 1-409
3. **Excel only**: Font formatting only applies to Excel files, not CSV

### Template Not Rendering

1. **Check syntax**: Ensure Jinja2 syntax is correct
2. **Check variables**: Variables must be wrapped in `{{ }}` or `{% %}`
3. **Test simple**: Try a simple template first: `{{ company }}`

### "Printed by" Row Not Showing

1. **Check setting**: Verify "Add 'Printed by' Row" is enabled
2. **Check template**: Template must be enabled and not empty

## Advanced Usage

### Dynamic Content Based on Report

```jinja2
Report: {{ report_name if report_name else doctype }}
{% if report_name %}
Report Type: Query Report
{% else %}
Report Type: List View
{% endif %}
```

### Date Formatting

```jinja2
Date: {{ date.strftime('%B %d, %Y') if date else '' }}
Time: {{ time.strftime('%I:%M %p') if time else '' }}
```

### Accessing User Information

```jinja2
User: {{ user_fullname }}
Email: {{ frappe.db.get_value("User", frappe.session.user, "email") }}
```

## Contributing

This app uses `pre-commit` for code formatting and linting. Please [install pre-commit](https://pre-commit.com/#installation) and enable it for this repository:

```bash
cd apps/export_letterhead
pre-commit install
```

Pre-commit is configured to use the following tools:
- ruff
- eslint
- prettier
- pyupgrade

## License

MIT

## Support

For issues, questions, or contributions, please visit the [GitHub repository](https://github.com/lijsamuael/export_letterhead/).

---

**Made with ❤️ for the Frappe/ERPNext community**
