// vybe-test: js/array_higher_order/partition_pattern
// origin: languages/js/tests/js/test_array_higher_order.rs

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

function partition(arr, pred) {
    return arr.reduce(([pass, fail], x) => {
        return pred(x) ? [[...pass, x], fail] : [pass, [...fail, x]];
    }, [[], []]);
}
const [evens, odds] = partition([1, 2, 3, 4, 5, 6], x => x % 2 === 0);
__check(__line(evens.join(",")), "2,4,6");
__check(__line(odds.join(",")), "1,3,5");
