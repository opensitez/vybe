// vybe-test: js/operators_deep/in_operator_considers_prototype_chain
// origin: languages/js/tests/js/test_operators_deep.rs

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

const proto = { inherited: 1 };
const child = Object.create(proto);
__check(__line("inherited" in child), "true");
__check(__line(child.hasOwnProperty("inherited")), "false");
