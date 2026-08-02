// vybe-test: js/try_catch_nested/nested_triple_null_throw_propagates
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
try{try{try{throw null;}catch(e){if(e===null){o.push("n");}else throw e;}}
catch{o.push("x");}}
catch{o.push("y");}
__check(__line(o.join(",")), "n");
