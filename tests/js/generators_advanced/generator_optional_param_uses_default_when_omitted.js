// vybe-test: js/generators_advanced/generator_optional_param_uses_default_when_omitted
// origin: languages/js/tests/js/test_generators_advanced.rs

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

function* range(start, end, step = 1) {
    for (let i = start; i < end; i += step) yield i;
}
console.log([...range(1, 6)].join(","));
console.log([...range(0, 10, 2)].join(","));
