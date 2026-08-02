// vybe-test: js/string_algorithms/string_rotate
// origin: languages/js/tests/js/test_string_algorithms.rs

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

const rotateStr = (s, n) => s.slice(n % s.length) + s.slice(0, n % s.length);
__check(__line(rotateStr("abcde", 2)), "cdeab");
__check(__line(rotateStr("hello", 0)), "hello");
__check(__line(rotateStr("abcde", 5)), "abcde");
