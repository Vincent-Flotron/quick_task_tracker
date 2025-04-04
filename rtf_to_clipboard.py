import win32clipboard

def copy_to_clipboard(rtf_content, plain_text_content):
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
    finally:
        # Close the clipboard
        win32clipboard.CloseClipboard()

# RTF content to copy
rtf_text = r"""{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
{\*\generator Riched20 10.0.18362;}viewkind4\uc1 
\pard\fs36\b Title of the Document\par
\fs24\i This is some italicized text.\par
\fs24 Here is a word that is \b bold\b0 .\par
}"""

# Plain text content to copy
plain_text = "Title of the Document\nThis is some italicized text.\nHere is a word that is bold."

# Copy the RTF and plain text content to the clipboard
copy_to_clipboard(rtf_text, plain_text)

print("RTF and plain text content copied to clipboard!")
