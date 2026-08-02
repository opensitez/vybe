// vybe-test: js/object_methods_deep/object_assign_only_copies_enumerable_own
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const src = Object.create({ inherited: true });
Object.defineProperty(src, "hidden", { value: 1, enumerable: false });
src.visible = 2;

const target = Object.assign({}, src);
__check(__line("inherited" in target), "false");
__check(__line("hidden" in target), "false");
__check(__line(target.visible), "2");
