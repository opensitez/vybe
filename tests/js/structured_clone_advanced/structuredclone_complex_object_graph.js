// vybe-test: js/structured_clone_advanced/structuredclone_complex_object_graph
// origin: languages/js/tests/js/test_structured_clone_advanced.rs

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

const orig = {
  users: [{ name: "Alice", scores: [1, 2, 3] }, { name: "Bob", scores: [4, 5] }],
  meta: { total: 2 }
};
const clone = structuredClone(orig);
clone.users[0].scores.push(99);
clone.meta.total = 99;
__check(__line(orig.users[0].scores.length), "3");
__check(__line(orig.meta.total), "2");
