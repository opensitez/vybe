// vybe-test: js/math_algorithms/fibonacci_sequence_generator
// origin: languages/js/tests/js/test_math_algorithms.rs

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

function* fibSeq() {
    let [a, b] = [0, 1];
    while (true) { yield a; [a, b] = [b, a+b]; }
}
const gen = fibSeq();
const first10 = Array.from({length: 10}, () => gen.next().value);
console.log(first10.join(","));
