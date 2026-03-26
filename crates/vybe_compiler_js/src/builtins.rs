/// Precompiled JS builtins — compiled to bytecode and injected into the VM.
/// These are regular JS functions that use the same opcodes as user code.
/// The compiler recognizes arr.map(fn) and routes to these.

/// Returns JS source code for all builtin functions.
/// These get compiled once at startup and installed as globals.
pub const BUILTINS_JS: &str = r#"
function __array_map(arr, fn) {
    let result = [];
    for (let i = 0; i < arr.length; i++) {
        result.push(fn(arr[i], i, arr));
    }
    return result;
}

function __array_filter(arr, fn) {
    let result = [];
    for (let i = 0; i < arr.length; i++) {
        if (fn(arr[i], i, arr)) {
            result.push(arr[i]);
        }
    }
    return result;
}

function __array_forEach(arr, fn) {
    for (let i = 0; i < arr.length; i++) {
        fn(arr[i], i, arr);
    }
}

function __array_find(arr, fn) {
    for (let i = 0; i < arr.length; i++) {
        if (fn(arr[i], i, arr)) {
            return arr[i];
        }
    }
    return null;
}

function __array_reduce(arr, fn, init) {
    let acc = init;
    for (let i = 0; i < arr.length; i++) {
        acc = fn(acc, arr[i], i, arr);
    }
    return acc;
}

function __array_sort(arr, fn) {
    let len = arr.length;
    for (let i = 0; i < len; i++) {
        for (let j = 0; j < len - i - 1; j++) {
            let cmp = 0;
            if (fn !== null) {
                cmp = fn(arr[j], arr[j + 1]);
            } else {
                if (arr[j] > arr[j + 1]) { cmp = 1; }
            }
            if (cmp > 0) {
                let temp = arr[j];
                arr[j] = arr[j + 1];
                arr[j + 1] = temp;
            }
        }
    }
    return arr;
}

function __array_some(arr, fn) {
    for (let i = 0; i < arr.length; i++) {
        if (fn(arr[i], i, arr)) { return true; }
    }
    return false;
}

function __array_every(arr, fn) {
    for (let i = 0; i < arr.length; i++) {
        if (!fn(arr[i], i, arr)) { return false; }
    }
    return true;
}

function __array_flat_map(arr, fn) {
    let result = [];
    for (let i = 0; i < arr.length; i++) {
        let mapped = fn(arr[i], i, arr);
        if (typeof mapped === "object") {
            for (let j = 0; j < mapped.length; j++) {
                result.push(mapped[j]);
            }
        } else {
            result.push(mapped);
        }
    }
    return result;
}

function __set_timeout(fn, ms) {
    clock.sleep(ms);
    fn();
}
"#;
