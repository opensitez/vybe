// vybe-test: js/functional_fp_patterns/continuation_passing_style
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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

function addCPS(a, b, k) { k(a + b); }
function multiplyCPS(a, b, k) { k(a * b); }
function sqrtCPS(n, k) { k(Math.sqrt(n)); }

addCPS(3, 4, sum =>
    multiplyCPS(sum, 2, product =>
        sqrtCPS(product, result =>
            __check(__line(Math.round(result * 100) / 100), "3.74")
        )
    )
);
