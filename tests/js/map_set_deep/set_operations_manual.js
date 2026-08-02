// vybe-test: js/map_set_deep/set_operations_manual
// origin: languages/js/tests/js/test_map_set_deep.rs

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

// Union
const union = new Set([...a, ...b]);
__check(__line([...union].sort((x,y) => x-y).join(",")), "1,2,3,4,5,6");

// Intersection
const intersection = new Set([...a].filter(x => b.has(x)));
__check(__line([...intersection].join(",")), "3,4");

// Difference (a - b)
const diff = new Set([...a].filter(x => !b.has(x)));
__check(__line([...diff].join(",")), "1,2");
