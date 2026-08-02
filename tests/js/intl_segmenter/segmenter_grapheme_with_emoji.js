// vybe-test: js/intl_segmenter/segmenter_grapheme_with_emoji
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
const text = "\uD83D\uDE00\uD83D\uDE01"; // two emoji
const segments = [...seg.segment(text)];
__check(__line(segments.length), "2");
