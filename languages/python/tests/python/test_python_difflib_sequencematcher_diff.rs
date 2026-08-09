use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Category 4: Sequence Comparison & Diffing (difflib module)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_difflib_sequencematcher_ratio() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "hello world", "hello word")
ratio = sm.ratio()
print(ratio > 0.9)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_difflib_get_close_matches() {
    let out = run_python(
        r#"
import difflib
words = ["apple", "banana", "apricot", "avocado"]
matches = difflib.get_close_matches("appl", words)
print(matches)
"#,
    );
    assert_eq!(out, vec!["['apple', 'apricot']"]);
}

#[test]
fn test_difflib_unified_diff() {
    let out = run_python(
        r#"
import difflib
a = ["line 1\n", "line 2\n"]
b = ["line 1\n", "line 2 modified\n"]
diff = list(difflib.unified_diff(a, b, fromfile="a.txt", tofile="b.txt"))
print(diff[0].strip())
print(diff[1].strip())
"#,
    );
    assert_eq!(out, vec!["--- a.txt", "+++ b.txt"]);
}

#[test]
fn test_difflib_context_diff() {
    let out = run_python(
        r#"
import difflib
a = ["one\n", "two\n"]
b = ["one\n", "three\n"]
diff = list(difflib.context_diff(a, b))
print(len(diff) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_difflib_differ_compare() {
    let out = run_python(
        r#"
import difflib
d = difflib.Differ()
res = list(d.compare(["a", "b"], ["a", "c"]))
print([line.strip() for line in res])
"#,
    );
    assert_eq!(out, vec!["['a', '- b', '+ c']"]);
}

#[test]
fn test_difflib_opcodes() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "abc", "axc")
opcodes = sm.get_opcodes()
tag_names = [op[0] for op in opcodes]
print("equal" in tag_names)
print("replace" in tag_names)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_difflib_matching_blocks() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "abcdef", "abxgef")
blocks = sm.get_matching_blocks()
print(len(blocks) >= 2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_difflib_isjunk_callback() {
    let out = run_python(
        r#"
import difflib
# Ignore spaces as junk
is_space = lambda x: x == " "
sm = difflib.SequenceMatcher(is_space, "a b c", "a  b c")
print(sm.ratio() > 0.8)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_difflib_get_close_matches_cutoff() {
    let out = run_python(
        r#"
import difflib
words = ["apple", "banana"]
matches = difflib.get_close_matches("xyz", words, cutoff=0.9)
print(matches)
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_difflib_get_close_matches_n_limit() {
    let out = run_python(
        r#"
import difflib
words = ["apple1", "apple2", "apple3", "apple4"]
matches = difflib.get_close_matches("apple", words, n=2)
print(len(matches))
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_difflib_quick_ratio() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "identical", "identical")
print(sm.quick_ratio())
print(sm.real_quick_ratio())
"#,
    );
    assert_eq!(out, vec!["1.0", "1.0"]);
}

#[test]
fn test_difflib_empty_sequences() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "", "")
print(sm.ratio())
"#,
    );
    assert_eq!(out, vec!["1.0"]);
}

#[test]
fn test_difflib_set_seqs() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher()
sm.set_seqs("first", "second")
print(sm.a)
print(sm.b)
"#,
    );
    assert_eq!(out, vec!["first", "second"]);
}

#[test]
fn test_difflib_set_seq1_set_seq2() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher()
sm.set_seq1("abc")
sm.set_seq2("abc")
print(sm.ratio())
"#,
    );
    assert_eq!(out, vec!["1.0"]);
}

#[test]
fn test_difflib_is_line_junk() {
    let out = run_python(
        r##"
import difflib
print(difflib.IS_LINE_JUNK("# comment"))
print(difflib.IS_LINE_JUNK("valid line"))
"##,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_difflib_is_character_junk() {
    let out = run_python(
        r#"
import difflib
print(difflib.IS_CHARACTER_JUNK(" "))
print(difflib.IS_CHARACTER_JUNK("\t"))
print(difflib.IS_CHARACTER_JUNK("a"))
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_difflib_restore_from_diff() {
    let out = run_python(
        r#"
import difflib
d = difflib.Differ()
diff = list(d.compare(["line 1"], ["line 2"]))
restored_1 = list(difflib.restore(diff, 1))
restored_2 = list(difflib.restore(diff, 2))
print(restored_1)
print(restored_2)
"#,
    );
    assert_eq!(out, vec!["['line 1']", "['line 2']"]);
}

#[test]
fn test_difflib_find_longest_match() {
    let out = run_python(
        r#"
import difflib
sm = difflib.SequenceMatcher(None, "xxxABCyyy", "zzzABCqqq")
match = sm.find_longest_match(0, 9, 0, 9)
print(f"a={match.a}, b={match.b}, size={match.size}")
"#,
    );
    assert_eq!(out, vec!["a=3, b=3, size=3"]);
}

#[test]
fn test_difflib_list_of_strings_comparison() {
    let out = run_python(
        r#"
import difflib
a = ["red", "blue", "green"]
b = ["red", "yellow", "green"]
sm = difflib.SequenceMatcher(None, a, b)
print(sm.ratio() > 0.6)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_difflib_htmldiff_class() {
    let out = run_python(
        r#"
import difflib
hd = difflib.HtmlDiff()
table = hd.make_table(["line 1"], ["line 2"])
print("<table" in table)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
