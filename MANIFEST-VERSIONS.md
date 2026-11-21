# Send to Wunderbuild: manifest register

The manifest tells Outlook where the add-in lives and what it may do. It is the file IT installs.
IT should install only a file whose SHA-256 matches the latest row below.
To check a file on Windows (PowerShell): `Get-FileHash manifest.xml -Algorithm SHA256`

Rules for changes: the add-in ID never changes; Version goes up with every change; permissions stay at ReadWriteItem unless approved by Jason in writing; every new row records the new hash.

| Manifest version | Date issued | SHA-256 | What changed |
|---|---|---|---|
| 1.0.0.20 | 21 Nov 2026 | `8f7285ff46f02eca175e3a3bb81ff963de528c8e4fa7fc1ca44885f55b8e04f4` | Send-time script (launchevent) 1.31: the handler for Archicad review emails (`#wb-<code>` in the subject) now looks for the tag again, every 0.3 s for up to 3 s, when the subject is still empty at the moment Outlook raises the event (a `mailto:` link can fill the subject in a moment later), and keeps a short log (no subjects or keys) that the Troubleshooter shows. Only the script address (?v=1.31) and the version text changed; permissions unchanged (ReadWriteItem). 1.0.0.19 was never given to IT. |
| 1.0.0.19 | 11 Nov 2026 | `2050fb5fdd5fa5bd8ed9376b12f4558a3d5edee68eb4c26b50489ca74a5594bb` | Send-time script (launchevent) 1.30: the "save to Wunderbuild?" prompt is now stricter. A bare 4 or 5 digit number no longer triggers it (needs job/project/#/No.); a one-word job name needs a second clue; J- references, multi-word job names and previously saved threads still ask. Only the script address (?v=1.30) changed; permissions unchanged (ReadWriteItem). 1.0.0.18 was never given to IT. Amended 19 Nov 2026 (still not given to IT): the button hover text now ends with the manifest version. |
| 1.0.0.18 | 5 Nov 2026 | `140ed780db0c638123dbe4e78e3d8e3808289915d53c495e08b4fd2837ed65af` | Send-time script (launchevent) 1.29: the "save to Wunderbuild?" prompt now appears only when the subject has a J- reference, a 4 or 5 digit job number or a job name that matches, or the thread was saved before from this computer. Only the script address (?v=1.29) changed; permissions unchanged (ReadWriteItem). 1.0.0.17 was never given to IT. |
| 1.0.0.17 | 4 Nov 2026 | `f4e349329703aaa59299fbb473f576dd83a73125230aa2f3f2eb0e7f57108048` | Worker address changed in the data domains (wb-proxy.tallwood.workers.dev, was wb-proxy.jason-984.workers.dev). Nothing else changed; permissions unchanged (ReadWriteItem). |
| 1.0.0.16 | 3 Nov 2026 | `2242a9ceea989d23357bc0b4110dc7939d4f9e495140f6ff8f0114d8ff0c1800` | Description updated; data domains now list the Wunderbuild preview site and the Worker. Permissions unchanged (ReadWriteItem). |
| 1.0.0.15 | earlier | `4afeabc4f3782f2acdd405c82a1736c6b7c41eb2c243be6c8719c56cd74be7c0` | Previous issue (not deployed). |
