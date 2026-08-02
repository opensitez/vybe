// vybe-test: js/math_algorithms/statistics_mean_median_mode
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

function mean(arr) { return arr.reduce((a,b)=>a+b,0) / arr.length; }
function median(arr) {
    const s = [...arr].sort((a,b)=>a-b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m-1]+s[m]) / 2;
}
function mode(arr) {
    const freq = new Map();
    for (const x of arr) freq.set(x, (freq.get(x)??0)+1);
    return [...freq.entries()].sort((a,b)=>b[1]-a[1])[0][0];
}
const data = [1, 2, 2, 3, 4, 4, 4, 5];
console.log(mean(data));
console.log(median(data));
console.log(mode(data));
