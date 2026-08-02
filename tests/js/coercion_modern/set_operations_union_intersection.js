// vybe-test: js/coercion_modern/set_operations_union_intersection
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let a = new Set([1, 2, 3, 4]);
let b = new Set([3, 4, 5, 6]);
let union = new Set([...a, ...b]);
let intersection = new Set([...a].filter(x => b.has(x)));
let difference = new Set([...a].filter(x => !b.has(x)));
__check(__line([...union].sort().join(",")), "1,2,3,4,5,6");
__check(__line([...intersection].sort().join(",")), "3,4");
__check(__line([...difference].sort().join(",")), "1,2");
