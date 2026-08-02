// vybe-test: js/generator_delegation_advanced/generator_chunk
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

function* chunk(iter, size) {
    let batch = [];
    for (const item of iter) {
        batch.push(item);
        if (batch.length === size) { yield batch; batch = []; }
    }
    if (batch.length) yield batch;
}
const chunks = [...chunk([1,2,3,4,5,6,7], 3)];
console.log(chunks.length);
console.log(chunks[0].join(","));
console.log(chunks[2].join(","));
