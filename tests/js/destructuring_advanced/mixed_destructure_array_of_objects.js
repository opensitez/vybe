// vybe-test: js/destructuring_advanced/mixed_destructure_array_of_objects
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

const [{ name: n1 }, { name: n2 }] = [{ name: "A" }, { name: "B" }];
__check(__line(n1, n2), "A B");
