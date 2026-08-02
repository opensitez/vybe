// vybe-test: js/try_catch_nested/sibling_finally_between_tries
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
try{o.push("t1");}finally{o.push("f");}
try{throw 0;}catch{o.push("c");}
__check(__line(o.join(",")), "t1,f,c");
