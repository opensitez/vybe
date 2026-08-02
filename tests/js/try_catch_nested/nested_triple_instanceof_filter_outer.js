// vybe-test: js/try_catch_nested/nested_triple_instanceof_filter_outer
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
try{try{try{throw new Error("x");}catch(e){throw e;}}
catch(e){if(e instanceof Error)o.push("e");else o.push("?");}}
catch{o.push("no");}
__check(__line(o.join(",")), "e");
