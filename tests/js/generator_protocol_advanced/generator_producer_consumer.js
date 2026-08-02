// vybe-test: js/generator_protocol_advanced/generator_producer_consumer
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* producer() {
    const items = [1, 2, 3, 4, 5];
    for (const item of items) {
        const doubled = yield item;
        if (doubled !== undefined) {
            // consumer sent back a value
        }
    }
}

const gen = producer();
const results = [];
let next = gen.next();
while (!next.done) {
    results.push(next.value * 2);
    next = gen.next();
}
console.log(results.join(","));
