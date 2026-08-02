// vybe-test: js/string_algorithms/zigzag_string_pattern
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

function zigzag(s, rows) {
    if (rows === 1) return s;
    const buckets = Array.from({length: rows}, () => "");
    let row = 0, dir = 1;
    for (const c of s) {
        buckets[row] += c;
        if (row === 0) dir = 1;
        else if (row === rows - 1) dir = -1;
        row += dir;
    }
    return buckets.join("");
}
console.log(zigzag("PAYPALISHIRING", 3));
console.log(zigzag("AB", 1));
