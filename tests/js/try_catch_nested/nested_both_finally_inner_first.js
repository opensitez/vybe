// vybe-test: js/try_catch_nested/nested_both_finally_inner_first
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
try{try{o.push("t");}finally{o.push("fi");}}
finally{o.push("fo");}
__check(__line(o.join(",")), "t,fi,fo");
