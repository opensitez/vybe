// vybe-test: js/generator_delegation_advanced/generator_zip_two
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

function* zip(...iters) {
    const gen = iters.map(i => i[Symbol.iterator]());
    while (true) {
        const results = gen.map(g => g.next());
        if (results.some(r => r.done)) break;
        yield results.map(r => r.value);
    }
}
const pairs = [...zip([1,2,3], ["a","b","c"])];
console.log(pairs.map(p => p.join(":")).join(","));
