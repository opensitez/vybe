// vybe-test: js/map_set_edge_matrix/set_entries_yields_value_value_pairs
// origin: languages/js/tests/js/test_map_set_edge_matrix.rs

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

const s = new Set([1, 2]);
const out = [];
for (const [a, b] of s.entries()) out.push(a + ":" + b);
console.log(out.join(","));
