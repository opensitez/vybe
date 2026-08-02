// vybe-test: js/intl_segmenter/segmenter_word_splits_words
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

const seg = new Intl.Segmenter("en", { granularity: "word" });
const text = "Hello World";
const words = [...seg.segment(text)].filter(s => s.isWordLike);
__check(__line(words.length), "2");
__check(__line(words[0].segment), "Hello");
__check(__line(words[1].segment), "World");
