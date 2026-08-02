// vybe-test: js/class_constructor_errors/finally_in_constructor_runs_before_throw_escapes
// origin: languages/js/tests/js/test_class_constructor_errors.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

let o=[];try{class C{constructor(){try{throw 1;}finally{o.push("f");}}} new C();}catch{o.push("c");}__check(__line(o.join(",")), "f,c");
