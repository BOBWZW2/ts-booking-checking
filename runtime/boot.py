"""Load the deployment owner's pinned private application into a temporary directory."""
import base64
import hashlib
import hmac
import io
import os
from pathlib import Path
import tarfile
import tempfile

bundle = base64.b64decode(os.environ.pop("APPLICATION_ARCHIVE_B64"), validate=True)
if not hmac.compare_digest(hashlib.sha256(bundle).hexdigest(), os.environ.pop("APPLICATION_ARCHIVE_SHA256")):
    raise RuntimeError("Application archive checksum mismatch")
folder = Path(tempfile.mkdtemp(prefix="service-"))
with tarfile.open(fileobj=io.BytesIO(bundle), mode="r:gz") as archive:
    for item in archive.getmembers():
        if not item.isfile() or item.name != Path(item.name).name or not item.name.endswith(".py") or item.size > 512_000:
            raise RuntimeError("Invalid application archive")
        (folder / item.name).write_bytes(archive.extractfile(item).read())
os.chdir(folder)
os.execvp("uvicorn", ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", os.environ["PORT"], "--workers", "1", "--no-access-log"])
