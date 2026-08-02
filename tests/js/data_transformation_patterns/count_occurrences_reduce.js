// vybe-test: js/data_transformation_patterns/count_occurrences_reduce
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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

const words = ["apple", "banana", "apple", "cherry", "banana", "apple"];
const counts = words.reduce((acc, w) => {
    acc[w] = (acc[w] ?? 0) + 1;
    return acc;
}, {});
__check(__line(counts.apple), "3");
__check(__line(counts.banana), "2");
__check(__line(counts.cherry), "1");
