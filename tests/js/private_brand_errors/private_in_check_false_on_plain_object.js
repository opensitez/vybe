// vybe-test: js/private_brand_errors/private_in_check_false_on_plain_object
// origin: languages/js/tests/js/test_private_brand_errors.rs

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

class C{#token=1; static isC(v){return #token in v;}} __check(__line(C.isC({})), "false");
