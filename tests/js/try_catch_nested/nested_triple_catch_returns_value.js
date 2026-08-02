// vybe-test: js/try_catch_nested/nested_triple_catch_returns_value
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

function f(){let o=[];
try{try{throw 5;}catch(e){return "got:"+e;}}
catch{return "outer";}}
__check(__line(f()), "got:5");
