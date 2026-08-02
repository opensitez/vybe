// vybe-test: js/nullish_coalescing_and_optional_chaining_combinations/test_js_optional_chaining_private_field_access
// origin: languages/js/tests/js/test_js_nullish_coalescing_and_optional_chaining_combinations.rs

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

class Secret {
    #code = 1234;
    getCode(obj) {
        return obj?.#code;
    }
}
const s = new Secret();
__check(__line(s.getCode(s) + "|" + (s.getCode(null) === undefined)), "1234|true");
