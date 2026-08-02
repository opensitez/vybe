// vybe-test: js/property_enumeration/for_in_traverses_prototype_chain
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const proto = { inherited: "yes" };
const obj = Object.create(proto);
obj.own = "yes";
const found = {};
for (const k in obj) found[k] = true;
console.log(found.own);
console.log(found.inherited);
