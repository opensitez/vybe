// vybe-test: js/closures_functional/reduce_to_frequency_map
// origin: languages/js/tests/js/test_closures_functional.rs

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

let letters = "abracadabra".split("");
let freq = letters.reduce((acc, ch) => {
    acc[ch] = (acc[ch] || 0) + 1;
    return acc;
}, {});
__check(__line(freq.a), "5");
__check(__line(freq.b), "2");
__check(__line(freq.r), "2");
