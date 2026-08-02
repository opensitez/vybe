// vybe-test: js/try_catch_nested/sibling_try_both_throw_caught_separately
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
try{throw "a";}catch(e){o.push(e);}
try{throw "b";}catch(e){o.push(e);}
__check(__line(o.join(",")), "a,b");
