// vybe-test: js/json_deep/parse_invalid_json_throws_syntax_error
// origin: languages/js/tests/js/test_json_deep.rs

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

let threw = false;
try { JSON.parse("{bad json}"); } catch (e) { threw = e instanceof SyntaxError; }
__check(__line(threw), "true");
