// vybe-test: js/destructuring_comprehensive/swap_via_destructuring
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

let x = 1, y = 2;
[x, y] = [y, x];
__check(__line(x), "2");
__check(__line(y), "1");
let a = "hello", b = "world";
[a, b] = [b, a];
__check(__line(a), "world");
__check(__line(b), "hello");
