// vybe-test: js/control_flow_advanced/for_of_iterable_iterator_return_called_on_throw
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

let nextCount = 0;
let returnCount = 0;
const iterable = {
    [Symbol.iterator]() {
        return {
            next() {
                nextCount++;
                return nextCount === 1
                    ? { value: nextCount, done: false }
                    : { done: true };
            },
            return() {
                returnCount++;
                return { done: true };
            }
        };
    }
};

try {
    for (const value of iterable) {
        if (value === 1) {
            throw new Error("loop failure");
        }
    }
} catch (e) {
    console.log(e.message);
}
console.log(`${nextCount}:${returnCount}`);
