/// Intl.Segmenter — word, sentence, grapheme segmentation

use super::helpers::run_js;

#[test]
fn segmenter_grapheme_splits_characters() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const segments = [...seg.segment("hello")];
console.log(segments.length);
console.log(segments[0].segment);
"#), vec!["5", "h"]);
}

#[test]
fn segmenter_grapheme_with_emoji() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const text = "\uD83D\uDE00\uD83D\uDE01"; // two emoji
const segments = [...seg.segment(text)];
console.log(segments.length);
"#), vec!["2"]);
}

#[test]
fn segmenter_word_splits_words() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "word" });
const text = "Hello World";
const words = [...seg.segment(text)].filter(s => s.isWordLike);
console.log(words.length);
console.log(words[0].segment);
console.log(words[1].segment);
"#), vec!["2", "Hello", "World"]);
}

#[test]
fn segmenter_word_iswordlike_false_for_spaces() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "word" });
const all = [...seg.segment("hi there")];
const spaces = all.filter(s => !s.isWordLike);
console.log(spaces.length > 0);
"#), vec!["true"]);
}

#[test]
fn segmenter_sentence_splits_sentences() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "sentence" });
const text = "Hello. World. Foo.";
const sentences = [...seg.segment(text)];
console.log(sentences.length);
"#), vec!["3"]);
}

#[test]
fn segmenter_segment_has_index() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const segs = [...seg.segment("abc")];
console.log(segs[0].index);
console.log(segs[1].index);
console.log(segs[2].index);
"#), vec!["0", "1", "2"]);
}

#[test]
fn segmenter_contains_full_input() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const text = "test";
const segs = [...seg.segment(text)];
const rejoined = segs.map(s => s.segment).join("");
console.log(rejoined === text);
"#), vec!["true"]);
}

#[test]
fn segmenter_resolved_options() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "word" });
const opts = seg.resolvedOptions();
console.log(opts.granularity);
console.log(typeof opts.locale);
"#), vec!["word", "string"]);
}

#[test]
fn segmenter_empty_string() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en", { granularity: "grapheme" });
const segs = [...seg.segment("")];
console.log(segs.length);
"#), vec!["0"]);
}

#[test]
fn segmenter_default_granularity_is_grapheme() {
    assert_eq!(run_js(r#"
const seg = new Intl.Segmenter("en");
const opts = seg.resolvedOptions();
console.log(opts.granularity);
"#), vec!["grapheme"]);
}
