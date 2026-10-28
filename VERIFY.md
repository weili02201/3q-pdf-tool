# Verifying a 3Q PDF download

Everything below can be recomputed from the files themselves.

| What | Where |
|---|---|
| Checksums, source scan, binary scan (JSON) | https://getwold.com/integrity.json |
| The same, with the commands to redo it | https://getwold.com/verify |

## Current release

- `3Q-PDF-Tool.exe` — 41 902 470 bytes
- SHA-256 `30ed7059afc84a6cac24ff20584cfac85e58316bc6cb5f8ac77edea3d617edac`

The executable ships split across two zips (`3qpdf.part1.zip`, `3qpdf.part2.zip`) because of its
size. Join the parts in order, then hash:

```
copy /b 3Q-PDF-Tool.exe.part1 + 3Q-PDF-Tool.exe.part2 3Q-PDF-Tool.exe
certutil -hashfile 3Q-PDF-Tool.exe SHA256
```

`JOIN-AND-RUN.bat` checks the size only, not the hash. Run `certutil` yourself if you want the
stronger check.

## Source in the zip vs this repository

The five files under `source/` inside the zip are this repository's files. Two of them
(`pdf_core.py`, `pdf_tool.py`) were saved with CRLF line endings when the zip was built; normalise
`\r\n` to `\n` before comparing hashes. The other three are byte-identical.

## What the checks do not establish

A checksum proves the file you received is the file that is served. It does not prove the file is
harmless, and PyInstaller builds are not byte-reproducible, so no hash can prove this `.exe` was
compiled from exactly this source. Building your own copy with `build_exe.py` is the way around
that limit.

## Network behaviour

`pdf_cli.py`, `pdf_core.py`, `pdf_tool.py` and `build_exe.py` import no networking module at all.
`pdf_mcp.py` imports `socket` and `http.server` for an inbound MCP server that listens on
`127.0.0.1` by default; its two socket calls are a local interface lookup and a hostname resolve.
No outbound connection exists on any code path.
