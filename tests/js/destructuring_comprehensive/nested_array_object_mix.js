// vybe-test: js/destructuring_comprehensive/nested_array_object_mix
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

const data = { items: [{ id: 1, tags: ["a", "b"] }, { id: 2, tags: ["c"] }] };
const { items: [{ id: firstId, tags: [firstTag] }, { tags: [secondItemTag] }] } = data;
__check(__line(firstId), "1");
__check(__line(firstTag), "a");
__check(__line(secondItemTag), "c");
