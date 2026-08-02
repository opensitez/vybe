// vybe-test: js/data_transformation_patterns/group_by_property
// origin: languages/js/tests/js/test_data_transformation_patterns.rs

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
    { cat: "A", val: 1 }, { cat: "B", val: 2 },
    { cat: "A", val: 3 }, { cat: "B", val: 4 }, { cat: "C", val: 5 }
];
const grouped = items.reduce((acc, item) => {
    const key = item.cat;
    (acc[key] ??= []).push(item.val);
    return acc;
}, {});
__check(__line(grouped.A.join(",")), "1,3");
__check(__line(grouped.B.join(",")), "2,4");
__check(__line(grouped.C.join(",")), "5");
