// vybe-test: js/intl_segmenter/segmenter_grapheme_splits_characters
// origin: languages/js/tests/js/test_intl_segmenter.rs

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

const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const segments = [...seg.segment("hello")];
__check(__line(segments.length), "5");
__check(__line(segments[0].segment), "h");
