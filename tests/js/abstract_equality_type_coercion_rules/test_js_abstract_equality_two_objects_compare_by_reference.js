// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_two_objects_compare_by_reference
// origin: languages/js/tests/js/test_js_abstract_equality_type_coercion_rules.rs

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

const obj1 = { a: 1 };
const obj2 = { a: 1 };
const obj3 = obj1;
__check(__line(`${obj1 == obj2}:${obj1 == obj3}`), "false:true");
