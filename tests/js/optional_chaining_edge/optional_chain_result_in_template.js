// vybe-test: js/optional_chaining_edge/optional_chain_result_in_template
// origin: languages/js/tests/js/test_optional_chaining_edge.rs

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

const user = { name: "Bob" };
const absent = null;
__check(__line(`Hello ${user?.name}`), "Hello Bob");
__check(__line(`Hello ${absent?.name ?? "stranger"}`), "Hello stranger");
