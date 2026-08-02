// vybe-test: js/array_algorithms/zip_unzip
// origin: languages/js/tests/js/test_array_algorithms.rs

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

const zip = (...arrs) => arrs[0].map((_, i) => arrs.map(a => a[i]));
const unzip = arrs => arrs[0].map((_, i) => arrs.map(a => a[i]));
const zipped = zip([1, 2, 3], ["a", "b", "c"]);
__check(__line(zipped[0].join(",")), "1,a");
__check(__line(zipped[2].join(",")), "3,c");
