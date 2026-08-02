// vybe-test: js/prototype_chain_deep/for_in_with_hasown_filters_to_own_only
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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
obj.own1 = 2;
obj.own2 = 3;
const ownKeys = [];
for (const k in obj) {
    if (Object.hasOwn(obj, k)) ownKeys.push(k);
}
console.log(ownKeys.sort().join(","));
