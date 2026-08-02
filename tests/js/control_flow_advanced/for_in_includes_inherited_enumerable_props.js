// vybe-test: js/control_flow_advanced/for_in_includes_inherited_enumerable_props
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

const parent = { inherited: true };
const child = Object.create(parent);
child.own = true;
const keys = [];
for (const k in child) keys.push(k);
console.log(keys.includes("inherited"));
console.log(keys.includes("own"));
