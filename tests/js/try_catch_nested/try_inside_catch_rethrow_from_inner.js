// vybe-test: js/try_catch_nested/try_inside_catch_rethrow_from_inner
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
try{throw new TypeError("t");}
catch(e){try{throw new RangeError("r");}catch(x){o.push(x.name);}}
__check(__line(o.join(",")), "RangeError");
