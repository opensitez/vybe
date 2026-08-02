// vybe-test: js/operator_misc/conditional_expression_short_circuit_side_effects
// origin: languages/js/tests/js/test_operator_misc.rs

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

let count = 0;
const inc = () => ++count;
true ? inc() : inc();  // only left branch
false ? inc() : inc(); // only right branch
__check(__line(count), "2");    // 2
