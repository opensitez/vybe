// vybe-test: js/number_advanced/random_seeded_lcg
// origin: languages/js/tests/js/test_number_advanced.rs

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

// Simple LCG pseudo-random for deterministic tests
function lcg(seed) {
    let s = seed;
    return () => {
        s = (1664525 * s + 1013904223) & 0xFFFFFFFF;
        return (s >>> 0) / 0x100000000;
    };
}
const rand = lcg(42);
const vals = Array.from({length: 5}, () => rand() > 0 && rand() < 1);
console.log(vals.every(Boolean));
const r1 = lcg(42)();
const r2 = lcg(42)();
console.log(r1 === r2);
