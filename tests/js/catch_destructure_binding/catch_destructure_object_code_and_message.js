// vybe-test: js/catch_destructure_binding/catch_destructure_object_code_and_message
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

try{throw{code:404,message:"nf"};}catch({code,message}){__check(__line(code), "404");__check(__line(message), "nf");}
