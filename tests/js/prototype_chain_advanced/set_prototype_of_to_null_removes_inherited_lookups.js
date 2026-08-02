// vybe-test: js/prototype_chain_advanced/set_prototype_of_to_null_removes_inherited_lookups
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const proto = { shared: 99 };
const obj = Object.create(proto);
__check(__line(obj.shared), "99");
Object.setPrototypeOf(obj, null);
__check(__line(obj.shared), "undefined");
__check(__line(Object.getPrototypeOf(obj) === null), "true");
