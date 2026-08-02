// vybe-test: js/array_hole_length_basics/push_after_length_extension_appends_at_new_end
// origin: languages/js/tests/js/test_array_hole_length_basics.rs

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

const arr = [1];
arr.length = 3;
arr.push(4);
__check(__line(arr.length), "4");
__check(__line(arr[3]), "4");
__check(__line(Object.keys(arr).join(",")), "0,3");
