// vybe-test: js/object_groupby/map_groupby_basic
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

const items = [1, 2, 3, 4, 5];
const groups = Map.groupBy(items, n => n % 2 === 0 ? "even" : "odd");
__check(__line(groups instanceof Map), "true");
__check(__line(groups.get("even").join(",")), "2,4");
__check(__line(groups.get("odd").join(",")), "1,3,5");
