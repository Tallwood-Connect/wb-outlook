# Send to Wunderbuild: manifest register

The manifest tells Outlook where the add-in lives and what it may do. It is the file IT installs.
IT should install only a file whose SHA-256 matches the latest row below.
To check a file on Windows (PowerShell): `Get-FileHash manifest.xml -Algorithm SHA256`

Rules for changes: the add-in ID never changes; Version goes up with every change; permissions stay at ReadWriteItem unless approved by Jason in writing; every new row records the new hash.

| Manifest version | Date issued | SHA-256 | What changed |
|---|---|---|---|
| 1.0.0.17 | 4 Nov 2026 | `f4e349329703aaa59299fbb473f576dd83a73125230aa2f3f2eb0e7f57108048` | Worker address changed in the data domains (wb-proxy.tallwood.workers.dev, was wb-proxy.jason-984.workers.dev). Nothing else changed; permissions unchanged (ReadWriteItem). |
| 1.0.0.16 | 3 Nov 2026 | `2242a9ceea989d23357bc0b4110dc7939d4f9e495140f6ff8f0114d8ff0c1800` | Description updated; data domains now list the Wunderbuild preview site and the Worker. Permissions unchanged (ReadWriteItem). |
| 1.0.0.15 | earlier | `4afeabc4f3782f2acdd405c82a1736c6b7c41eb2c243be6c8719c56cd74be7c0` | Previous issue (not deployed). |
