// vybe-test: js/object_groupby/map_groupby_with_object_keys
// origin: languages/js/tests/js/test_object_groupby.rs

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

const keyA = { type: "A" };
const keyB = { type: "B" };
const items = [
    { val: 1, key: keyA },
    { val: 2, key: keyB },
    { val: 3, key: keyA }
];
const groups = Map.groupBy(items, item => item.key);
__check(__line(groups.get(keyA).length), "2");
__check(__line(groups.get(keyB).length), "1");
