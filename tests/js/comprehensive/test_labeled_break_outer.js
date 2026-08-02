// vybe-test: js/comprehensive/test_labeled_break_outer
// origin: languages/js/tests/js/js_comprehensive_test.rs

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

let result = 0;
        outer: for (let i = 0; i < 5; i++) {
            for (let j = 0; j < 5; j++) {
                if (j === 2) break outer;
                result++;
            }
        }
        console.log(result);
