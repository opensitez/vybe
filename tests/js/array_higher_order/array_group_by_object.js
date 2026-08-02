// vybe-test: js/array_higher_order/array_group_by_object
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

const items = [
    { type: "fruit", name: "apple" },
    { type: "veggie", name: "carrot" },
    { type: "fruit", name: "banana" },
];
const grouped = Object.groupBy(items, x => x.type);
__check(__line(grouped.fruit.length), "2");
__check(__line(grouped.veggie.length), "1");
__check(__line(grouped.fruit[0].name), "apple");
