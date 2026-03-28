use super::*;

pub fn register(vm: &mut VM) {
    register_datetime_ns(vm);
    register_stringbuilder_ns(vm);
    register_collections_ns(vm);
    register_timespan_ns(vm);
    register_guid_ns(vm);
    register_primitives_ns(vm);
    register_process_ns(vm);
    register_array_statics_ns(vm);
}

fn register_datetime_ns(vm: &mut VM) {
    // DateTime (direct)
    let dt = ensure_namespace(vm, &["DateTime"]);
    set_prop(&dt, "now", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&dt, "parse", host_fn_ref(vm, "vybe:types", "dateTimeParse"));
    set_prop(&dt, "today", host_fn_ref(vm, "vybe:types", "dateTimeNow"));

    // System.DateTime
    let sys_dt = ensure_namespace(vm, &["System", "DateTime"]);
    set_prop(&sys_dt, "now", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&sys_dt, "parse", host_fn_ref(vm, "vybe:types", "dateTimeParse"));
    set_prop(&sys_dt, "today", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&sys_dt, "utcnow", host_fn_ref(vm, "vybe:types", "dateTimeNow"));
    set_prop(&sys_dt, "maxvalue", Value::F64(253402300799.0)); // 9999-12-31
    set_prop(&sys_dt, "minvalue", Value::F64(0.0));
    set_prop(&sys_dt, "daysinmonth", host_fn_ref(vm, "vybe:types", "dateTimeNow")); // placeholder
    set_prop(&sys_dt, "isleapyear", host_fn_ref(vm, "vybe:types", "dateTimeNow")); // placeholder
}

fn register_stringbuilder_ns(vm: &mut VM) {
    let sb = ensure_namespace(vm, &["System", "Text", "StringBuilder"]);
    set_prop(&sb, "new", host_fn_ref(vm, "vybe:types", "stringBuilderNew"));
}

fn register_collections_ns(vm: &mut VM) {
    let coll = ensure_namespace(vm, &["System", "Collections", "Generic"]);

    // Namespace paths
    let queue_ns = ensure_namespace(vm, &["System", "Collections", "Generic", "Queue"]);
    set_prop(&queue_ns, "new", host_fn_ref(vm, "vybe:types", "queueNew"));
    let stack_ns = ensure_namespace(vm, &["System", "Collections", "Generic", "Stack"]);
    set_prop(&stack_ns, "new", host_fn_ref(vm, "vybe:types", "stackNew"));
    let hs_ns = ensure_namespace(vm, &["System", "Collections", "Generic", "HashSet"]);
    set_prop(&hs_ns, "new", host_fn_ref(vm, "vybe:types", "hashSetNew"));
    set_prop(&coll, "list", host_fn_ref(vm, "vybe:types", "listNew"));
    set_prop(&coll, "dictionary", host_fn_ref(vm, "vybe:types", "dictNew"));

    // Bare globals for `new List<T>()`, `new Dictionary<K,V>()`, etc.
    vm.globals.insert("list".into(), host_fn_ref(vm, "vybe:types", "listNew"));
    vm.globals.insert("dictionary".into(), host_fn_ref(vm, "vybe:types", "dictNew"));
    vm.globals.insert("queue".into(), host_fn_ref(vm, "vybe:types", "queueNew"));
    vm.globals.insert("stack".into(), host_fn_ref(vm, "vybe:types", "stackNew"));
    vm.globals.insert("hashset".into(), host_fn_ref(vm, "vybe:types", "hashSetNew"));
    vm.globals.insert("arraylist".into(), host_fn_ref(vm, "vybe:types", "listNew"));
    vm.globals.insert("hashtable".into(), host_fn_ref(vm, "vybe:types", "dictNew"));
    vm.globals.insert("stringbuilder".into(), host_fn_ref(vm, "vybe:types", "stringBuilderNew"));
}

fn register_timespan_ns(vm: &mut VM) {
    let ts = ensure_namespace(vm, &["TimeSpan"]);
    set_prop(&ts, "fromdays", host_fn_ref(vm, "vybe:types", "timeSpanfromDays"));
    set_prop(&ts, "fromhours", host_fn_ref(vm, "vybe:types", "timeSpanfromHours"));
    set_prop(&ts, "fromminutes", host_fn_ref(vm, "vybe:types", "timeSpanfromMinutes"));
    set_prop(&ts, "fromseconds", host_fn_ref(vm, "vybe:types", "timeSpanfromSeconds"));
    set_prop(&ts, "frommilliseconds", host_fn_ref(vm, "vybe:types", "timeSpanfromMilliseconds"));
    set_prop(&ts, "zero", host_fn_ref(vm, "vybe:types", "timeSpanZero"));

    let sys_ts = ensure_namespace(vm, &["System", "TimeSpan"]);
    set_prop(&sys_ts, "fromdays", host_fn_ref(vm, "vybe:types", "timeSpanfromDays"));
    set_prop(&sys_ts, "fromhours", host_fn_ref(vm, "vybe:types", "timeSpanfromHours"));
    set_prop(&sys_ts, "fromminutes", host_fn_ref(vm, "vybe:types", "timeSpanfromMinutes"));
    set_prop(&sys_ts, "fromseconds", host_fn_ref(vm, "vybe:types", "timeSpanfromSeconds"));
    set_prop(&sys_ts, "frommilliseconds", host_fn_ref(vm, "vybe:types", "timeSpanfromMilliseconds"));
    set_prop(&sys_ts, "zero", host_fn_ref(vm, "vybe:types", "timeSpanZero"));
}

fn register_guid_ns(vm: &mut VM) {
    let guid = ensure_namespace(vm, &["Guid"]);
    set_prop(&guid, "newguid", host_fn_ref(vm, "vybe:types", "guidNewGuid"));
    set_prop(&guid, "empty", host_fn_ref(vm, "vybe:types", "guidEmpty"));
    set_prop(&guid, "parse", host_fn_ref(vm, "vybe:types", "guidParse"));

    let sys_guid = ensure_namespace(vm, &["System", "Guid"]);
    set_prop(&sys_guid, "newguid", host_fn_ref(vm, "vybe:types", "guidNewGuid"));
    set_prop(&sys_guid, "empty", host_fn_ref(vm, "vybe:types", "guidEmpty"));
    set_prop(&sys_guid, "parse", host_fn_ref(vm, "vybe:types", "guidParse"));
}

fn register_primitives_ns(vm: &mut VM) {
    // System.Double
    let dbl = ensure_namespace(vm, &["System", "Double"]);
    set_prop(&dbl, "parse", host_fn_ref(vm, "vybe:types", "doubleParse"));
    set_prop(&dbl, "tryparse", host_fn_ref(vm, "vybe:types", "doubleTryParse"));
    set_prop(&dbl, "maxvalue", Value::F64(f64::MAX));
    set_prop(&dbl, "minvalue", Value::F64(f64::MIN));
    set_prop(&dbl, "nan", Value::F64(f64::NAN));
    set_prop(&dbl, "positiveinfinity", Value::F64(f64::INFINITY));
    set_prop(&dbl, "negativeinfinity", Value::F64(f64::NEG_INFINITY));

    // System.Single
    let sng = ensure_namespace(vm, &["System", "Single"]);
    set_prop(&sng, "parse", host_fn_ref(vm, "vybe:types", "doubleParse"));
    set_prop(&sng, "maxvalue", Value::F64(f32::MAX as f64));
    set_prop(&sng, "minvalue", Value::F64(f32::MIN as f64));

    // System.Boolean
    let bln = ensure_namespace(vm, &["System", "Boolean"]);
    set_prop(&bln, "parse", host_fn_ref(vm, "vybe:types", "booleanParse"));

    // System.Decimal (alias to Double for now)
    let dec = ensure_namespace(vm, &["System", "Decimal"]);
    set_prop(&dec, "parse", host_fn_ref(vm, "vybe:types", "doubleParse"));

    // System.DBNull
    let dbnull = ensure_namespace(vm, &["System", "DBNull"]);
    set_prop(&dbnull, "value", Value::Null);

    // System.EventArgs
    let ea = ensure_namespace(vm, &["System", "EventArgs"]);
    set_prop(&ea, "empty", Value::Null);
}

fn register_process_ns(vm: &mut VM) {
    let proc = ensure_namespace(vm, &["Process"]);
    set_prop(&proc, "start", host_fn_ref(vm, "vybe:types", "processStart"));

    let sys_proc = ensure_namespace(vm, &["System", "Diagnostics", "Process"]);
    set_prop(&sys_proc, "start", host_fn_ref(vm, "vybe:types", "processStart"));
}

fn register_array_statics_ns(vm: &mut VM) {
    let sys_arr = ensure_namespace(vm, &["System", "Array"]);
    set_prop(&sys_arr, "clear", host_fn_ref(vm, "vybe:types", "arrayClear"));
    set_prop(&sys_arr, "copy", host_fn_ref(vm, "vybe:types", "arrayCopy"));
    set_prop(&sys_arr, "resize", host_fn_ref(vm, "vybe:types", "arrayResize"));
    set_prop(&sys_arr, "sort", host_fn_ref(vm, "vybe:types", "arraySort"));
    set_prop(&sys_arr, "reverse", host_fn_ref(vm, "vybe:array", "reverse"));
    set_prop(&sys_arr, "indexof", host_fn_ref(vm, "vybe:array", "indexOf"));

    // System.Tuple
    let tuple = ensure_namespace(vm, &["System", "Tuple"]);
    set_prop(&tuple, "create", host_fn_ref(vm, "vybe:array", "from")); // simplified

    // System.BitConverter
    let bc = ensure_namespace(vm, &["System", "BitConverter"]);
    set_prop(&bc, "tostring", host_fn_ref(vm, "vybe:convert", "toString"));
    set_prop(&bc, "todouble", host_fn_ref(vm, "vybe:convert", "cdbl"));
}
