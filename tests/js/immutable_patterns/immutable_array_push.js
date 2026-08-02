// vybe-test: js/immutable_patterns/immutable_array_push
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

const arr = Object.freeze([1, 2, 3]);
const newArr = [...arr, 4];
__check(__line(arr.length), "3");
__check(__line(newArr.length), "4");
__check(__line(newArr.join(",")), "1,2,3,4");
