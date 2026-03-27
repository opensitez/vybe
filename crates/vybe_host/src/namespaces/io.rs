use super::*;

pub fn register(vm: &mut VM) {
    // File (direct shortcut)
    let file = ensure_namespace(vm, &["File"]);
    register_file_methods(vm, &file);

    // System.IO.File
    let sys_file = ensure_namespace(vm, &["System", "IO", "File"]);
    register_file_methods(vm, &sys_file);

    // System.IO.Path
    let sys_path = ensure_namespace(vm, &["System", "IO", "Path"]);
    set_prop(&sys_path, "combine", host_fn_ref(vm, "wasi:filesystem", "pathCombine"));
    set_prop(&sys_path, "getfilename", host_fn_ref(vm, "wasi:filesystem", "pathGetFileName"));
    set_prop(&sys_path, "getextension", host_fn_ref(vm, "wasi:filesystem", "pathGetExtension"));
    set_prop(&sys_path, "getdirectoryname", host_fn_ref(vm, "wasi:filesystem", "pathGetDirectory"));
    set_prop(&sys_path, "getfilenamewithoutextension", host_fn_ref(vm, "wasi:filesystem", "pathGetFileNameWithoutExt"));
    set_prop(&sys_path, "changeextension", host_fn_ref(vm, "wasi:filesystem", "pathChangeExtension"));
    set_prop(&sys_path, "getfullpath", host_fn_ref(vm, "wasi:filesystem", "pathGetFullPath"));
    set_prop(&sys_path, "gettemppath", host_fn_ref(vm, "wasi:filesystem", "pathGetTempPath"));

    // Directory (direct shortcut)
    let dir = ensure_namespace(vm, &["Directory"]);
    register_directory_methods(vm, &dir);

    // System.IO.Directory
    let sys_dir = ensure_namespace(vm, &["System", "IO", "Directory"]);
    register_directory_methods(vm, &sys_dir);

    // Path (direct shortcut)
    let path = ensure_namespace(vm, &["Path"]);
    set_prop(&path, "combine", host_fn_ref(vm, "wasi:filesystem", "pathCombine"));
    set_prop(&path, "getfilename", host_fn_ref(vm, "wasi:filesystem", "pathGetFileName"));
    set_prop(&path, "getextension", host_fn_ref(vm, "wasi:filesystem", "pathGetExtension"));
    set_prop(&path, "getdirectoryname", host_fn_ref(vm, "wasi:filesystem", "pathGetDirectory"));
    set_prop(&path, "getfilenamewithoutextension", host_fn_ref(vm, "wasi:filesystem", "pathGetFileNameWithoutExt"));
    set_prop(&path, "changeextension", host_fn_ref(vm, "wasi:filesystem", "pathChangeExtension"));
    set_prop(&path, "getfullpath", host_fn_ref(vm, "wasi:filesystem", "pathGetFullPath"));
    set_prop(&path, "gettemppath", host_fn_ref(vm, "wasi:filesystem", "pathGetTempPath"));

    // IO (shortcut namespace for IO.File, IO.Path, etc.)
    let io = ensure_namespace(vm, &["IO"]);
    set_prop(&io, "file", ensure_namespace(vm, &["File"]));
    set_prop(&io, "path", ensure_namespace(vm, &["Path"]));
}

fn register_file_methods(vm: &VM, ns: &Value) {
    set_prop(ns, "readalltext", host_fn_ref(vm, "wasi:filesystem", "readFile"));
    set_prop(ns, "writealltext", host_fn_ref(vm, "wasi:filesystem", "writeFile"));
    set_prop(ns, "appendalltext", host_fn_ref(vm, "wasi:filesystem", "appendFile"));
    set_prop(ns, "exists", host_fn_ref(vm, "wasi:filesystem", "exists"));
    set_prop(ns, "delete", host_fn_ref(vm, "wasi:filesystem", "remove"));
    set_prop(ns, "copy", host_fn_ref(vm, "wasi:filesystem", "copy"));
    set_prop(ns, "move", host_fn_ref(vm, "wasi:filesystem", "rename"));
}

fn register_directory_methods(vm: &VM, ns: &Value) {
    set_prop(ns, "exists", host_fn_ref(vm, "wasi:filesystem", "isDir"));
    set_prop(ns, "createdirectory", host_fn_ref(vm, "wasi:filesystem", "mkdir"));
    set_prop(ns, "delete", host_fn_ref(vm, "wasi:filesystem", "remove"));
    set_prop(ns, "getfiles", host_fn_ref(vm, "wasi:filesystem", "listDir"));
    set_prop(ns, "getcurrentdirectory", host_fn_ref(vm, "wasi:cli", "cwd"));
}
