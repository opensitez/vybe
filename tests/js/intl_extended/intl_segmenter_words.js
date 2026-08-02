// vybe-test: js/intl_extended/intl_segmenter_words
// origin: languages/js/tests/js/test_intl_extended.rs

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

const seg = new Intl.Segmenter("en-US", { granularity: "word" });
const segments = [...seg.segment("Hello world")];
const wordSegments = segments.filter(s => s.isWordLike);
__check(__line(wordSegments.length), "2");
