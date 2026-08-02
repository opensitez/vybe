// vybe-test: js/try_catch_nested/nested_inner_rethrows_error
// origin: languages/js/tests/js/test_try_catch_nested.rs

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

let o=[];
try{try{throw new Error("a");}catch(e){o.push("inner");throw e;}}
catch(e){o.push("outer:"+e.message);}
__check(__line(o.join(",")), "inner,outer:a");
