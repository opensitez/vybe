// vybe-test: js/urlsearchparams/urlsearchparams_constructs_from_object_record
// origin: languages/js/tests/js/test_urlsearchparams.rs

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

const params = new URLSearchParams({ lang: "js", level: "advanced" });
__check(__line(params.get("lang")), "js");
__check(__line(params.get("level")), "advanced");
