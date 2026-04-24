// node:path — Node-style path module for Vybe.
//
// Re-exports path primitives from `wasi:filesystem`. Node API shape
// (`path.join`, `path.basename`, `path.extname`) requires local
// translation in the adapter; today we expose the underlying names.

export { pathCombine, pathGetDirectory, pathGetExtension, pathGetFileName, pathGetFileNameWithoutExt, pathChangeExtension, pathHasExtension, pathIsRooted, pathGetTempPath, pathGetFullPath } from "wasi:filesystem";
