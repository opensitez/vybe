// vybe-test: js/generator_advanced_patterns/generator_coroutine_via_send
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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

function* accumulator() {
    let sum = 0;
    while (true) {
        const n = yield sum;
        if (n === null) break;
        sum += n;
    }
}
const gen = accumulator();
gen.next();       // start
gen.next(10);
gen.next(20);
const result = gen.next(5);
console.log(result.value); // 35
