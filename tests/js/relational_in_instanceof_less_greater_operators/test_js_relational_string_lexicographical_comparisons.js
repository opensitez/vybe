// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_relational_string_lexicographical_comparisons
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

__check(__line(`${"apple" < "banana"}:${"2" > "10"}:${"a" <= "a"}`), "true:true:true");
