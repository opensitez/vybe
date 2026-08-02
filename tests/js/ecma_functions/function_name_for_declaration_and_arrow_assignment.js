// vybe-test: js/ecma_functions/function_name_for_declaration_and_arrow_assignment
// origin: languages/js/tests/js/test_ecma_functions.rs

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

function greet() {}
const answer = () => 42;
__check(__line(greet.name), "greet");
__check(__line(answer.name), "answer");
