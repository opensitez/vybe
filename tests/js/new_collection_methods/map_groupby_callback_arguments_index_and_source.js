// vybe-test: js/new_collection_methods/map_groupby_callback_arguments_index_and_source
// origin: languages/js/tests/js/test_new_collection_methods.rs

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

const out = []; Map.groupBy(["a", "b"], (val, idx) => { out.push(idx + ":" + val); return idx; }); __check(__line(out.join(",")), "0:a,1:b");
