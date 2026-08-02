// vybe-test: js/generators/generator_infinite_sequence
// origin: languages/js/tests/js/test_generators.rs

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

function* naturals() {
    let n = 1;
    while (true) {
        yield n++;
    }
}
let gen = naturals();
let results = [];
for (let i = 0; i < 5; i++) {
    results.push(gen.next().value);
}
console.log(results.join(","));
