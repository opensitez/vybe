// vybe-test: js/try_catch_nested/sibling_try_finally_independent
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
try{o.push("1");}finally{o.push("f1");}
try{o.push("2");}finally{o.push("f2");}
__check(__line(o.join(",")), "1,f1,2,f2");
