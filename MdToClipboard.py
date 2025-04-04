import win32clipboard
import html2text
import markdown


class md_to_clipboard:
    def copy_to_clipboard(rtf_content, plain_text_content, html_content):
        status = ""

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
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text_content)

            # Register the HTML format
            html_format = win32clipboard.RegisterClipboardFormat("HTML Format")

            # Set the HTML content to the clipboard
            win32clipboard.SetClipboardData(html_format, html_content.encode('utf-8'))
            
        except Exception as e:
            raise Exception(f"Error setting clipboard data: {e}")
        finally:
            # Close the clipboard
            win32clipboard.CloseClipboard()

        return True
    


    def create_html_with_fragment(html_body):
        html_header = "Version:1.0\nStartHTML:{0:010d}\nEndHTML:{1:010d}\nStartFragment:{2:010d}\nEndFragment:{3:010d}\n"
        start_html = len(html_header.format(0, 0, 0, 0).encode('utf-8'))
        end_html = start_html + len(html_body.encode('utf-8'))
        start_fragment = start_html
        end_fragment = end_html
        html_content = html_header.format(start_html, end_html, start_fragment, end_fragment) + html_body
        return html_content

    def html_to_rtf(html_body):
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html_body, 'html.parser')
        rtf_content = r"""{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
    {\*\generator Riched20 10.0.18362;}viewkind4\uc1
    """

        for tag in soup.descendants:
            if tag.name == 'p':
                rtf_content += r"\pard "
                if tag.strong:
                    rtf_content += r"\b "
                if tag.em:
                    rtf_content += r"\i "
                rtf_content += tag.get_text().replace('\n', ' ') + r" \par "

        rtf_content += r"}"
        return rtf_content


    def md_to_clipboard_for_onenote(markdown_text):
        """"RTF, plain text, and HTML content copied to clipboard!"""

        # Convert Markdown to HTML
        html_body = f"<html><body>{markdown.markdown(markdown_text)}</body></html>"

        html_text = md_to_clipboard.create_html_with_fragment(html_body)

        # Generate plain text from HTML
        plain_text = html2text.html2text(html_body)

        # Generate RTF text from HTML
        rtf_text = md_to_clipboard.html_to_rtf(html_body)

        # Copy the RTF, plain text, and HTML content to the clipboard
        if md_to_clipboard.copy_to_clipboard(rtf_text, plain_text, html_text):
            return True
        else:
            return False


if __name__ == "__main__":

    markdown_text = """
    - (Dev) Soply
        - TaskN°bla
            - booking : [https://bla.ch](https://bla.ch)
            - delivered
                - [] Testy
                - [x] Testc
    """

if md_to_clipboard.md_to_clipboard_for_onenote(markdown_text):
    print("RTF, plain text, and HTML content copied to clipboard!")

