// vybe-test: js/symbol_registry_matrix/symbol_registry_value_can_index_object_consistently
// origin: languages/js/tests/js/test_symbol_registry_matrix.rs

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

const s1 = Symbol.for("reg");
const s2 = Symbol.for("reg");
const obj = { [s1]: 99 };
console.log(obj[s2]);
