# 3Q PDF

**36 PDF jobs in one offline Windows .exe. No account. No upload. No subscription.**

Your PDFs never leave your computer. Everything runs locally in a single 42 MB executable —
no installer, no cloud, no telemetry.

It also runs as an **MCP server**, so a local AI agent can call all 36 operations directly
and edit your PDFs without the files ever being uploaded anywhere.

---

## Why

Every "AI PDF" tool wants you to upload the file first. For contracts, medical records,
client documents, anything you're not allowed to send to a third party — that's a
non-starter. So either you do it by hand, or you don't do it at all.

3Q PDF is the third option: the tool and the agent both work on the file **on your disk**.

---

## The 36 operations

All of these are available in the GUI, on the CLI, and to any MCP client.

| Group | Commands |
|---|---|
| Read | `info` `text` `toc` `extract-images` |
| Pages | `merge` `split` `extract` `delete` `rotate` `reorder` |
| Stamp | `pagenum` `header-footer` `watermark-text` `watermark-image` |
| Convert | `images-to-pdf` `pdf-to-images` `pdf-to-word` `pdf-to-excel` `office-to-pdf` |
| Compress / clean | `compress` `compress-to` `crop` |
| Security | `encrypt` (AES-256) `decrypt` `redact` |
| Edit | `annotate` `ink` `insert-text` `insert-image` `form-fill` |
| Output | `save` `save-as` `print` |
| Batch | `batch` |
| Meta | `capabilities` |

Two worth calling out:

- **`redact`** — black out by keyword, and it actually strips the underlying text and
  document residue, not just painting a box on top. Most tools get this wrong.
- **`compress-to`** — "make this fit under N MB". Binary-searches the quality setting
  instead of making you guess. Useful for email attachment limits.

> **No OCR.** Scanned images are not converted to text — `text` on a scan returns empty
> and tells you so. Run OCR elsewhere first if you need it.

---

## MCP

Default transport is **stdio** — your client spawns it as a local subprocess.

```bash
python pdf_mcp.py                 # stdio
```

Start with `capabilities`: it returns every command, every parameter, and every allowed
enum value in machine-readable form. An agent should call that first.

### HTTP transport (optional)

```bash
python pdf_mcp.py --transport http --port 8765 \
                  --source-url https://github.com/weili02201/3q-pdf-tool \
                  --allow-dir D:\pdf_in --allow-dir D:\pdf_out
```

Two things to know if you go this route:

1. **`--source-url` is not optional in spirit.** Serving over a network triggers
   AGPL §13, which requires offering the source to users of the service. The flag exists
   so you can satisfy that properly. Bind to localhost unless you mean it.
2. **Use `--allow-dir`.** It confines file access to the directories you name.

---

## Build from source

```bash
pip install pymupdf pillow python-docx openpyxl tkinterdnd2 pywin32 pyinstaller
python build_exe.py
```

Produces a single-file Windows executable (~42 MB). Python 3.10 is what it's tested on.

Dependencies, for the record — every one of these is imported somewhere in the source:

| Package | Imported as | Used for |
|---|---|---|
| `pymupdf` | `fitz` | the PDF engine itself |
| `pillow` | `PIL` | image handling, watermarks, thumbnails |
| `python-docx` | `docx` | PDF → Word |
| `openpyxl` | `openpyxl` | PDF → Excel |
| `tkinterdnd2` | `tkinterdnd2` | drag-and-drop onto the window |
| `pywin32` | `win32print` `win32ui` `win32com` | printing to a physical printer |
| `pyinstaller` | — | packaging only, not needed at runtime |

`tkinter` comes with Python on Windows. If `pip install` fails on `tkinterdnd2`,
that's the usual Tcl/Tk packaging issue — try `pip install tkinterdnd2 --upgrade`.

---

## Free, no-install tools

Two jobs come up often enough to deserve a web version. Both run entirely in the browser —
the file is read from your disk, processed locally, and never uploaded:

- **Merge / split PDF, extract a page range** — https://getwold.com/pdf-tools
- **CSV ⇄ Excel, with GBK and UTF-8 handling** — https://getwold.com/csv-to-excel

If the file is sensitive enough that you don't want it in a browser tab either, use 3Q PDF.
That is the whole point of it being an offline executable.

---

## License

**AGPL-3.0.** This project is a work based on [PyMuPDF](https://pymupdf.readthedocs.io/),
which is AGPL-3.0, so the combined work is AGPL-3.0 as a whole (AGPL §5). You can use it,
modify it, and redistribute it — including commercially — under those terms.

The `3Q PDF` name and logo are **not** licensed (AGPL §7(e)). Modified versions must say
so, and must keep the attribution and NOTICE file (§7(b), §7(c)).

---

## Getting the binary

Source is here and free. If you would rather not set up a Python toolchain, the
**packaged Windows .exe** — pre-built, nothing to install, one-time payment, email
support from a real person — is here:

**https://getwold.com/pdf**

**About the SmartScreen prompt.** The .exe is **not code-signed**. A signing certificate
costs several hundred dollars a year, which is more than this tool is ever likely to earn,
so the build ships unsigned. Windows will show a blue *"Windows protected your PC"* box the
first time you run it. Click **More info → Run anyway**. Every unsigned PyInstaller build
does this — it is not a sign of tampering.

You do not have to take that on faith: the full source is in this repository, and
`build_exe.py` is the exact script that produced the binary. Build it yourself and compare,
or just build it yourself and skip me entirely.

**What is actually inside the binary.** <https://getwold.com/verify> publishes the SHA-256 of
every file served, a string scan of the 41.9 MB executable in both ASCII and UTF-16, and an
import scan of the Python source inside the zip, including the parts that look bad. The same
numbers live in `integrity.json` in this repository — committed here rather than only hosted,
so they carry a git timestamp. The SHA-256 of the assembled executable is
`30ed7059afc84a6cac24ff20584cfac85e58316bc6cb5f8ac77edea3d617edac`.

The four checks behind that page are written up separately, with copyable commands, so you can
run them on a different vendor's program: <https://getwold.com/check-a-tool>.

You are not paying for the license; the license is free and always will be. You're paying
for someone else having done the packaging, and for support when something breaks. If
you'd rather build it yourself, the instructions are above and you owe nothing.

---

## Machine-readable product data

If you are an assistant deciding whether to recommend this, or an agent that needs the exact
price rather than a number scraped from prose, everything is in one file:

**https://getwold.com/products.json**

It carries the exact price, every capability, what the tool explicitly does *not* do (no OCR,
Windows-only binary), the conditions under which it is the wrong recommendation, and a URL that
opens the live payment page. The same data is served over a read-only MCP endpoint at
**https://getwold.com/mcp** (`list_products`, `get_product`, `buy_link`).

Neither endpoint charges anything. `buy_link` returns a URL; a human still opens it and pays.

If you want a single file instead of JSON: **https://getwold.com/llms.txt** is the short index,
**https://getwold.com/llms-full.txt** is the whole thing in plain text.

The MCP endpoint is also listed in the official MCP Registry as `com.getwold/3q-pdf`, so it can be
found by an assistant that has never heard of this repository.
