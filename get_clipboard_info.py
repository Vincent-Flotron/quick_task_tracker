import win32clipboard

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

if __name__ == "__main__":
    clipboard_info = get_clipboard_info()
    display_clipboard_info(clipboard_info)
