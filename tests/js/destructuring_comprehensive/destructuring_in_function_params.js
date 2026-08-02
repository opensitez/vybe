// vybe-test: js/destructuring_comprehensive/destructuring_in_function_params
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

function point({ x = 0, y = 0, z = 0 } = {}) {
    return `${x},${y},${z}`;
}
__check(__line(point({ x: 1, y: 2 })), "1,2,0");
__check(__line(point({ z: 5 })), "0,0,5");
__check(__line(point()), "0,0,0");
