/** Trigger a browser download for a Blob or byte array. */
export function downloadBlob(data: Blob | Uint8Array, fileName: string, mime?: string): void {
  const blob =
    data instanceof Blob
      ? data
      : new Blob([data as BlobPart], { type: mime ?? "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  // Revoke on the next tick so the download has a chance to start.
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
