// vybe-test: js/set_operations_es2025/chain_union_then_intersection
// origin: languages/js/tests/js/test_set_operations_es2025.rs

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

const a = new Set([1, 2, 3]);
const b = new Set([3, 4, 5]);
const c = new Set([2, 3, 6]);
const result = a.union(b).intersection(c);
__check(__line([...result].sort((a,b)=>a-b).join(",")), "2,3");
