// vybe-test: js/string_algorithms/roman_to_int
// origin: languages/js/tests/js/test_string_algorithms.rs

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

function romanToInt(s) {
    const map = { I:1, V:5, X:10, L:50, C:100, D:500, M:1000 };
    let result = 0;
    for (let i = 0; i < s.length; i++) {
        const curr = map[s[i]], next = map[s[i+1]];
        result += (next > curr) ? -curr : curr;
    }
    return result;
}
console.log(romanToInt("III"));
console.log(romanToInt("IV"));
console.log(romanToInt("MCMXCIV"));
