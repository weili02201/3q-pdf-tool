# Code signing policy

**What is signed.** Only the Windows executable built by `build_exe.py` from the source in this
repository, and only when the build is produced by the repository itself rather than on a
contributor's machine.

**Who signs.** The maintainer of this repository (@weili02201), as the sole author and approver. No
change to this repository is merged without review by that same maintainer, and every signing
request is approved manually.

**How to verify.** The SHA-256 of each release is published at https://getwold.com/integrity.json
and recorded in this repository. The current release is not code-signed; Windows SmartScreen will
warn on first run, and https://getwold.com/verify explains how to check the binary by hash instead.

**What would invalidate a signature.** A signed binary that does not hash to the value recorded in
this repository should be treated as tampered with.

This policy exists so the signing arrangement is documented rather than implied. It is a statement
of current practice, not a promise that a certificate exists today.
