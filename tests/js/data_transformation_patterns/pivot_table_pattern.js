// vybe-test: js/data_transformation_patterns/pivot_table_pattern
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

const sales = [
    { region: "North", product: "A", amount: 100 },
    { region: "South", product: "A", amount: 200 },
    { region: "North", product: "B", amount: 150 },
    { region: "South", product: "B", amount: 250 },
];
const pivot = sales.reduce((acc, s) => {
    if (!acc[s.region]) acc[s.region] = {};
    acc[s.region][s.product] = (acc[s.region][s.product] ?? 0) + s.amount;
    return acc;
}, {});
__check(__line(pivot.North.A), "100");
__check(__line(pivot.South.B), "250");
