// vybe-test: js/catch_destructure_binding/catch_destructure_nested_array_in_object
// origin: languages/js/tests/js/test_catch_destructure_binding.rs

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

try{throw{pair:[3,4]};}catch({pair:[x,y]}){__check(__line(x+y), "7");}
