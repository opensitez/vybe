// vybe-test: js/object_groupby/object_groupby_returns_null_prototype_object
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

const groups = Object.groupBy([1, 2, 3], n => "key");
__check(__line("key" in groups), "true");
__check(__line(groups.key.length), "3");
