// vybe-test: js/map_set_deep_patterns/map_to_object_and_back
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

const obj = { a: 1, b: 2, c: 3 };
const map = new Map(Object.entries(obj));
map.set("d", 4);
const back = Object.fromEntries(map);
__check(__line(back.a), "1");
__check(__line(back.d), "4");
__check(__line(Object.keys(back).sort().join(",")), "a,b,c,d");
