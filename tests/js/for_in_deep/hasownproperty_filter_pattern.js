// vybe-test: js/for_in_deep/hasownproperty_filter_pattern
// origin: languages/js/tests/js/test_for_in_deep.rs

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
const obj = Object.create(proto);
obj.own = 2;
const ownKeys = [];
for (const k in obj) {
    if (Object.hasOwn(obj, k)) ownKeys.push(k);
}
console.log(ownKeys.join(","));
