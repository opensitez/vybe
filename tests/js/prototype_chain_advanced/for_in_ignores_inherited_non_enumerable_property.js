// vybe-test: js/prototype_chain_advanced/for_in_ignores_inherited_non_enumerable_property
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

const proto = {};
Object.defineProperty(proto, "hidden", {
    value: 99,
    enumerable: false
});
const obj = Object.create(proto);
obj.visible = true;

const keys = [];
for (const key in obj) {
    keys.push(key);
}

console.log(keys.join(","));
console.log("hidden" in obj);
