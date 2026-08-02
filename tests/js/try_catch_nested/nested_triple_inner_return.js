// vybe-test: js/try_catch_nested/nested_triple_inner_return
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
try{try{try{return 1;}catch{return 2;}}
catch{return 3;}}
catch{return 4;}}
__check(__line(f()), "1");
