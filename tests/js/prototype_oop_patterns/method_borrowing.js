// vybe-test: js/prototype_oop_patterns/method_borrowing
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

const arrayLike = { 0: "a", 1: "b", 2: "c", length: 3 };
const joined = Array.prototype.join.call(arrayLike, "-");
const mapped = Array.prototype.map.call(arrayLike, s => s.toUpperCase());
__check(__line(joined), "a-b-c");
__check(__line(mapped.join(",")), "A,B,C");
