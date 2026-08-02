// vybe-test: js/arrow_functions/arrow_to_string_contains_arrow
// origin: languages/js/tests/js/test_arrow_functions.rs

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

__check(__line(Function.prototype.toString.call(()=>1).includes("=>")), "true");
