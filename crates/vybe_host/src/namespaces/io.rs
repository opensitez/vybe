use super::*;

pub fn register(vm: &mut VM) {
    // JS `fs.*` ambient compatibility namespace.
    let fs = ensure_namespace(vm, &["fs"]);
    for name in &[
        "readFile",
        "writeFile",
        "appendFile",
        "exists",
        "isFile",
        "isDir",
        "remove",
        "listDir",
        "mkdir",
        "fileSize",
        "rename",
        "copy",
        "stat",
        "readDirEntries",
    ] {
        set_prop(&fs, name, host_fn_ref(vm, "wasi:filesystem", name));
    }
}
