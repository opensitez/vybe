// vybe-test: js/string_prototype_methods/repeat_throws_on_negative
// origin: languages/js/tests/js/test_string_prototype_methods.rs

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

try{"a".repeat(-1); console.log("ok");}catch(e){console.log(e instanceof RangeError);}
