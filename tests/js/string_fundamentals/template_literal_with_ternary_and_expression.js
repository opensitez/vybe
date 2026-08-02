// vybe-test: js/string_fundamentals/template_literal_with_ternary_and_expression
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

const user = { first: "Ada", last: "Lovelace", active: true };
__check(__line(`${user.first} ${user.last} is ${user.active ? "active" : "inactive"}`), "Ada Lovelace is active");
const total = 3 + 4;
__check(__line(`total=${total}`), "total=7");
