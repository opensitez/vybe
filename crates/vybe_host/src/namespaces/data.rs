use super::*;

pub fn register(vm: &mut VM) {
    // System.Data.DataTable
    let dt = ensure_namespace(vm, &["System", "Data", "DataTable"]);
    set_prop(&dt, "new", host_fn_ref(vm, "vybe:data", "dataTableNew"));

    // System.Data.DataSet
    let ds = ensure_namespace(vm, &["System", "Data", "DataSet"]);
    set_prop(&ds, "new", host_fn_ref(vm, "vybe:data", "dataSetNew"));

    // System.Drawing.Point / Size / Font
    let pt = ensure_namespace(vm, &["System", "Drawing", "Point"]);
    set_prop(&pt, "new", host_fn_ref(vm, "vybe:drawing", "pointNew"));

    let sz = ensure_namespace(vm, &["System", "Drawing", "Size"]);
    set_prop(&sz, "new", host_fn_ref(vm, "vybe:drawing", "sizeNew"));

    let font = ensure_namespace(vm, &["System", "Drawing", "Font"]);
    set_prop(&font, "new", host_fn_ref(vm, "vybe:drawing", "fontNew"));

    // Direct shortcuts — namespace objects with .new
    let pt_s = ensure_namespace(vm, &["Point"]);
    set_prop(&pt_s, "new", host_fn_ref(vm, "vybe:drawing", "pointNew"));

    let sz_s = ensure_namespace(vm, &["Size"]);
    set_prop(&sz_s, "new", host_fn_ref(vm, "vybe:drawing", "sizeNew"));

    let font_s = ensure_namespace(vm, &["Font"]);
    set_prop(&font_s, "new", host_fn_ref(vm, "vybe:drawing", "fontNew"));

    // Also register as bare callable globals for `new Point(10, 20)` pattern
    vm.globals.insert("point".into(), host_fn_ref(vm, "vybe:drawing", "pointNew"));
    vm.globals.insert("size".into(), host_fn_ref(vm, "vybe:drawing", "sizeNew"));
    vm.globals.insert("sizef".into(), host_fn_ref(vm, "vybe:drawing", "sizeNew"));
    vm.globals.insert("font".into(), host_fn_ref(vm, "vybe:drawing", "fontNew"));
}
