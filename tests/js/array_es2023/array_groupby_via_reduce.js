// vybe-test: js/array_es2023/array_groupby_via_reduce
// origin: languages/js/tests/js/test_array_es2023.rs

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

const items = ["apple", "avocado", "banana", "blueberry", "cherry"];
const grouped = items.reduce((acc, item) => {
    const key = item[0];
    if (!acc[key]) acc[key] = [];
    acc[key].push(item);
    return acc;
}, {});
__check(__line(grouped["a"].length), "2");
__check(__line(grouped["b"].length), "2");
__check(__line(grouped["c"].length), "1");
