import win32clipboard

def copy_to_clipboard(rtf_content, plain_text_content, html_content):
    # Open the clipboard
    win32clipboard.OpenClipboard()
    try:
        # Clear the clipboard
        win32clipboard.EmptyClipboard()

        # Register the RTF format
        rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")

        # Set the RTF content to the clipboard
        win32clipboard.SetClipboardData(rtf_format, rtf_content.encode('utf-8'))

        # Set the plain text content to the clipboard
        win32clipboard.SetClipboardData(win32clipboard.CF_TEXT, plain_text_content.encode('utf-8'))

        # Register the HTML format
        html_format = win32clipboard.RegisterClipboardFormat("HTML Format")

        # Set the HTML content to the clipboard
        win32clipboard.SetClipboardData(html_format, html_content.encode('utf-8'))

        print("Clipboard set successfully.")
    except Exception as e:
        print(f"Error setting clipboard data: {e}")
    finally:
        # Close the clipboard
        win32clipboard.CloseClipboard()

def get_clipboard_info():
    # Open the clipboard
    win32clipboard.OpenClipboard()
    try:
        # Get the available clipboard formats
        formats = win32clipboard.EnumClipboardFormats(0)
        clipboard_info = {}

        # Iterate through available formats
        while formats:
            format_type = formats
            try:
                # Try to get the data for the current format
                data = win32clipboard.GetClipboardData(format_type)
                clipboard_info[format_type] = {
                    'data': data,
                    'type': format_type
                }
            except Exception as e:
                clipboard_info[format_type] = {
                    'error': str(e)
                }
            formats = win32clipboard.EnumClipboardFormats(formats)

        return clipboard_info
    finally:
        # Close the clipboard
        win32clipboard.CloseClipboard()

def display_clipboard_info(info):
    for format_type, content in info.items():
        print(f"Format: {format_type}")
        if 'data' in content:
            print(f"Data: {content['data']}")
        if 'error' in content:
            print(f"Error: {content['error']}")
        print("-" * 40)

def create_html_with_fragment(html_body):
    html_header = "Version:1.0\nStartHTML:{0:010d}\nEndHTML:{1:010d}\nStartFragment:{2:010d}\nEndFragment:{3:010d}\n"
    start_html = len(html_header.format(0, 0, 0, 0))
    end_html = start_html + len(html_body)
    start_fragment = start_html
    end_fragment = end_html
    html_content = html_header.format(start_html, end_html, start_fragment, end_fragment) + html_body
    return html_content

# RTF content to copy
rtf_text = r"""{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
{\*\generator Riched20 10.0.18362;}viewkind4\uc1
\pard\fs36\b Title of the Document\par
\fs24\i This is some italicized text.\par
\fs24 Here is a word that is \b bold\b0 .\par
}"""

# Plain text content to copy
plain_text = "Title of the Document\nThis is some italicized text.\nHere is a word that is bold."

# HTML content to copy
html_body = """<html><body>
<p><strong>Title of the Document</strong></p>
<p><em>This is some italicized text.</em></p>
<p>Here is a word that is <strong>bold</strong>.</p>
</body></html>"""

html_text = create_html_with_fragment(html_body)

# Copy the RTF, plain text, and HTML content to the clipboard
copy_to_clipboard(rtf_text, plain_text, html_text)

# Get and display clipboard info
clipboard_info = get_clipboard_info()
display_clipboard_info(clipboard_info)

print("RTF, plain text, and HTML content copied to clipboard!")
