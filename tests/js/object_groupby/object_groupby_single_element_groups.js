// vybe-test: js/object_groupby/object_groupby_single_element_groups
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

const words = ["apple", "banana", "cherry"];
const groups = Object.groupBy(words, w => w[0]);
__check(__line(groups.a[0]), "apple");
__check(__line(groups.b[0]), "banana");
__check(__line(groups.c[0]), "cherry");
