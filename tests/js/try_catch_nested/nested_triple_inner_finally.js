// vybe-test: js/try_catch_nested/nested_triple_inner_finally
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
try{try{try{throw 1;}finally{o.push("f3");}}
catch(e){o.push("c2");}}
catch(e){o.push("c1");}
__check(__line(o.join(",")), "f3,c2");
