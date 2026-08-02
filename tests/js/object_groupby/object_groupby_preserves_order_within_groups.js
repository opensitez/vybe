// vybe-test: js/object_groupby/object_groupby_preserves_order_within_groups
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

const letters = ["c", "a", "b", "a", "c", "b"];
const groups = Object.groupBy(letters, l => l);
__check(__line(groups.a.join(",")), "a,a");
__check(__line(groups.b.join(",")), "b,b");
__check(__line(groups.c.join(",")), "c,c");
