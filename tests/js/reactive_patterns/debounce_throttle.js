// vybe-test: js/reactive_patterns/debounce_throttle
// origin: languages/js/tests/js/test_reactive_patterns.rs

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

function debounce(fn, delay) {
    let timer;
    return (...args) => {
        clearTimeout(timer);
        timer = setTimeout(() => fn(...args), delay);
    };
}
function throttle(fn, limit) {
    let lastTime = 0;
    return (...args) => {
        const now = Date.now();
        if (now - lastTime >= limit) { lastTime = now; fn(...args); }
    };
}
// Verify they return functions
console.log(typeof debounce(() => {}, 100));
console.log(typeof throttle(() => {}, 100));
const calls = [];
const t = throttle(x => calls.push(x), 1000);
t(1); t(2); t(3);
console.log(calls.length);
