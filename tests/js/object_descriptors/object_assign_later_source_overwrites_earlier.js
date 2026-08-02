// vybe-test: js/object_descriptors/object_assign_later_source_overwrites_earlier
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const result = Object.assign({}, { x: 1 }, { x: 2 }, { x: 3 });
__check(__line(result.x), "3");
