// vybe-test: js/map_set_edge_matrix/map_for_each_receives_this_arg
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

const m = new Map([["a", 1], ["b", 2]]);
const ctx = { prefix: ">" };
const out = [];
m.forEach(function(value, key) { out.push(this.prefix + key + value); }, ctx);
console.log(out.join(","));
