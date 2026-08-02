// vybe-test: js/try_catch_nested/nested_number_throw_caught_as_number
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

let n=0;
try{try{throw 42;}catch(e){n=e;}}
catch{n=-1;}
__check(__line(n), "42");
