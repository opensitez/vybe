// vybe-test: js/try_catch_nested/nested_triple_finally_after_inner_catch
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
try{try{try{throw 0;}catch{o.push("c");}}
finally{o.push("f2");}}
finally{o.push("f1");}
__check(__line(o.join(",")), "c,f2,f1");
