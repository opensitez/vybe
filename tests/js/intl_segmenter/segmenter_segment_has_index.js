// vybe-test: js/intl_segmenter/segmenter_segment_has_index
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
const segs = [...seg.segment("abc")];
__check(__line(segs[0].index), "0");
__check(__line(segs[1].index), "1");
__check(__line(segs[2].index), "2");
