// vybe-test: js/try_catch_nested/nested_triple_throw_caught_at_level_one
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
try{try{try{throw new RangeError("r");}catch(e){throw e;}}
catch(e){throw e;}}
catch(e){o.push(e.name);}
__check(__line(o.join(",")), "RangeError");
