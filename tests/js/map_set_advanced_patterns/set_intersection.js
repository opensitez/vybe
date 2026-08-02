// vybe-test: js/map_set_advanced_patterns/set_intersection
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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
const b = new Set([2, 4, 6]);
const intersection = new Set([...a].filter(x => b.has(x)));
__check(__line([...intersection].sort((a,b)=>a-b).join(",")), "2,4");
