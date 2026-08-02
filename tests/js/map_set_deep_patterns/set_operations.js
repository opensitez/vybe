// vybe-test: js/map_set_deep_patterns/set_operations
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

const a = new Set([1, 2, 3, 4]);
const b = new Set([3, 4, 5, 6]);
const union = new Set([...a, ...b]);
const intersection = new Set([...a].filter(x => b.has(x)));
const difference = new Set([...a].filter(x => !b.has(x)));
__check(__line([...union].sort((a,b)=>a-b).join(",")), "1,2,3,4,5,6");
__check(__line([...intersection].join(",")), "3,4");
__check(__line([...difference].join(",")), "1,2");
