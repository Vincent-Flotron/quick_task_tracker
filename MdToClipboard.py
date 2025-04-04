import win32clipboard
import html2text
import markdown
from bs4 import BeautifulSoup
from io import StringIO


class MdToClipboard:
    @staticmethod
    def copy_to_clipboard(rtf_content: str, plain_text: str, html_content: str) -> bool:
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()

            # Set RTF
            rtf_format = win32clipboard.RegisterClipboardFormat("Rich Text Format")
            win32clipboard.SetClipboardData(rtf_format, rtf_content.encode('utf-8'))

            # Set plain text
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)

            # Set HTML
            html_format = win32clipboard.RegisterClipboardFormat("HTML Format")
            win32clipboard.SetClipboardData(html_format, html_content.encode('utf-8'))

        except Exception as e:
            raise RuntimeError(f"Clipboard operation failed: {e}")

        finally:
            win32clipboard.CloseClipboard()

        return True

    @staticmethod
    def create_html_with_fragment(html_body: str) -> str:
        # Placeholder header with fixed size
        header_template = (
            "Version:1.0\n"
            "StartHTML:{0:010d}\n"
            "EndHTML:{1:010d}\n"
            "StartFragment:{2:010d}\n"
            "EndFragment:{3:010d}\n"
        )
        header_size = len(header_template.format(0, 0, 0, 0).encode('utf-8'))

        start_html = header_size
        end_html = start_html + len(html_body.encode('utf-8'))

        html = header_template.format(start_html, end_html, start_html, end_html) + html_body
        return html

    @staticmethod
    def html_to_rtf(html_body: str) -> str:
        soup = BeautifulSoup(html_body, 'html.parser')
        rtf = StringIO()
        rtf.write(r"""{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0 Calibri;}}""")
        rtf.write(r"{\*\generator Python;}viewkind4\uc1\pard ")

        for tag in soup.find_all(['p']):
            rtf.write(r"\pard ")

            if tag.find('strong'):
                rtf.write(r"\b ")
            if tag.find('em'):
                rtf.write(r"\i ")

            rtf.write(tag.get_text().replace('\n', ' '))
            rtf.write(r" \par ")

        rtf.write(r"}")
        return rtf.getvalue()

    @staticmethod
    def md_to_clipboard_for_onenote(markdown_text: str) -> bool:
        html_body = f"<html><body>{markdown.markdown(markdown_text)}</body></html>"

        html_text = MdToClipboard.create_html_with_fragment(html_body)
        print("*************")
        print(html_text)
        plain_text = html2text.html2text(html_body)
        print("*************")
        print(plain_text)
        rtf_text = MdToClipboard.html_to_rtf(html_body)

        return MdToClipboard.copy_to_clipboard(rtf_text, plain_text, html_text)

    @staticmethod
    def html_to_clipboard_for_onenote(html_body: str) -> bool:
        html_body = f"<html><body>{html_body}</body></html>"

        html_text = MdToClipboard.create_html_with_fragment(html_body)
        print("*************")
        print(html_text)
        plain_text = html2text.html2text(html_body)
        print("*************")
        print(plain_text)
        rtf_text = MdToClipboard.html_to_rtf(html_body)

        return MdToClipboard.copy_to_clipboard(rtf_text, plain_text, html_text)

if __name__ == "__main__":
    markdown_text = """
- (Dev) Soply
    - TaskN°bla
        - booking : [https://bla.ch](https://bla.ch)
        - delivered
            - [] Testy
            - [x] Testc
"""

    if MdToClipboard.md_to_clipboard_for_onenote(markdown_text):
        print("RTF, plain text, and HTML content copied to clipboard!")
