// vybe-test: js/string_processing_deep/slug_generation
// origin: languages/js/tests/js/test_string_processing_deep.rs

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

function slugify(text) {
    return text.toLowerCase()
        .replace(/[^\w\s-]/g, "")
        .replace(/[\s_]+/g, "-")
        .replace(/^-+|-+$/g, "");
}
__check(__line(slugify("Hello World!")), "hello-world");
__check(__line(slugify("  The Quick Brown Fox  ")), "the-quick-brown-fox");
__check(__line(slugify("Hello---World")), "hello---world");
