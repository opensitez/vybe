// vybe-test: js/generator_state_machines/token_lexer_generator
// origin: languages/js/tests/js/test_generator_state_machines.rs

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

function* tokenize(expr) {
    const re = /\d+|[+\-*/()]/g;
    let m;
    while ((m = re.exec(expr)) !== null) {
        yield m[0];
    }
}
const tokens = [...tokenize("1 + 2 * (3 - 4)")];
console.log(tokens.join(","));
