export type TFileFingerprint = {
  name: string;
  size: number;
  lastModified: number;
};

export const fingerprintFile = (file: File): TFileFingerprint => ({
  name: file.name,
  size: file.size,
  lastModified: file.lastModified,
});

export const isSameFingerprint = (a: TFileFingerprint, b: TFileFingerprint) => a.name === b.name && a.size === b.size && a.lastModified === b.lastModified;
