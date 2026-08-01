use crate::helpers::{run_print, run_python_one};

#[test]
fn list_slice_start_stop() {
    assert_eq!(
        run_python_one("x = [0, 1, 2, 3, 4]\nprint(x[1:4])\n"),
        "[1, 2, 3]"
    );
}

#[test]
fn list_slice_from_start() {
    assert_eq!(
        run_python_one("x = [10, 20, 30]\nprint(x[:2])\n"),
        "[10, 20]"
    );
}

#[test]
fn list_slice_to_end() {
    assert_eq!(
        run_python_one("x = [10, 20, 30]\nprint(x[1:])\n"),
        "[20, 30]"
    );
}

#[test]
fn list_slice_full_copy() {
    assert_eq!(run_python_one("x = [1, 2, 3]\nprint(x[:])\n"), "[1, 2, 3]");
}

#[test]
fn list_slice_negative_start() {
    assert_eq!(
        run_python_one("x = [1, 2, 3, 4]\nprint(x[-3:-1])\n"),
        "[2, 3]"
    );
}

#[test]
fn list_slice_negative_stop() {
    assert_eq!(
        run_python_one("x = [1, 2, 3, 4]\nprint(x[1:-1])\n"),
        "[2, 3]"
    );
}

#[test]
fn list_slice_step_two() {
    assert_eq!(
        run_python_one("x = [0, 1, 2, 3, 4, 5]\nprint(x[::2])\n"),
        "[0, 2, 4]"
    );
}

#[test]
fn list_slice_reverse() {
    assert_eq!(
        run_python_one("x = [1, 2, 3]\nprint(x[::-1])\n"),
        "[3, 2, 1]"
    );
}

#[test]
fn list_slice_reverse_with_bounds() {
    assert_eq!(
        run_python_one("x = [0, 1, 2, 3, 4]\nprint(x[3:0:-1])\n"),
        "[3, 2, 1]"
    );
}

#[test]
fn list_index_positive() {
    assert_eq!(run_python_one("x = [5, 6, 7]\nprint(x[1])\n"), "6");
}

#[test]
fn list_index_negative_last() {
    assert_eq!(run_python_one("x = [5, 6, 7]\nprint(x[-1])\n"), "7");
}

#[test]
fn list_index_negative_first() {
    assert_eq!(run_python_one("x = [5, 6, 7]\nprint(x[-3])\n"), "5");
}

#[test]
fn str_slice_start_stop() {
    assert_eq!(run_print("'abcdef'[1:4]"), "bcd");
}

#[test]
fn str_slice_from_start() {
    assert_eq!(run_print("'abcdef'[:3]"), "abc");
}

#[test]
fn str_slice_to_end() {
    assert_eq!(run_print("'abcdef'[3:]"), "def");
}

#[test]
fn str_slice_full() {
    assert_eq!(run_print("'hi'[:]"), "hi");
}

#[test]
fn str_slice_negative_indices() {
    assert_eq!(run_print("'python'[-4:-1]"), "tho");
}

#[test]
fn str_slice_step_two() {
    assert_eq!(run_print("'abcdef'[::2]"), "ace");
}

#[test]
fn str_slice_reverse() {
    assert_eq!(run_print("'stressed'[::-1]"), "desserts");
}

#[test]
fn str_index_positive() {
    assert_eq!(run_print("'abc'[0]"), "a");
}

#[test]
fn str_index_negative() {
    assert_eq!(run_print("'abc'[-1]"), "c");
}

#[test]
fn tuple_slice_start_stop() {
    assert_eq!(
        run_python_one("t = (1, 2, 3, 4)\nprint(t[1:3])\n"),
        "(2, 3)"
    );
}

#[test]
fn tuple_slice_from_start() {
    assert_eq!(
        run_python_one("t = (10, 20, 30)\nprint(t[:2])\n"),
        "(10, 20)"
    );
}

#[test]
fn tuple_slice_to_end() {
    assert_eq!(run_python_one("t = (10, 20, 30)\nprint(t[2:])\n"), "(30,)");
}

#[test]
fn tuple_slice_full() {
    assert_eq!(run_python_one("t = (1, 2)\nprint(t[:])\n"), "(1, 2)");
}

#[test]
fn tuple_slice_negative() {
    assert_eq!(
        run_python_one("t = (1, 2, 3, 4)\nprint(t[-2:])\n"),
        "(3, 4)"
    );
}

#[test]
fn tuple_slice_step() {
    assert_eq!(
        run_python_one("t = (0, 1, 2, 3, 4)\nprint(t[::2])\n"),
        "(0, 2, 4)"
    );
}

#[test]
fn tuple_slice_reverse() {
    assert_eq!(
        run_python_one("t = (1, 2, 3)\nprint(t[::-1])\n"),
        "(3, 2, 1)"
    );
}

#[test]
fn tuple_index_positive() {
    assert_eq!(run_python_one("t = (9, 8, 7)\nprint(t[0])\n"), "9");
}

#[test]
fn tuple_index_negative() {
    assert_eq!(run_python_one("t = (9, 8, 7)\nprint(t[-2])\n"), "8");
}

#[test]
fn list_slice_empty_range() {
    assert_eq!(run_python_one("x = [1, 2, 3]\nprint(x[2:2])\n"), "[]");
}

#[test]
fn list_slice_beyond_bounds() {
    assert_eq!(run_python_one("x = [1, 2, 3]\nprint(x[1:99])\n"), "[2, 3]");
}

#[test]
fn str_slice_empty_range() {
    assert_eq!(run_print("'hello'[3:3]"), "");
}

#[test]
fn list_slice_step_three() {
    assert_eq!(
        run_python_one("x = [0, 1, 2, 3, 4, 5, 6]\nprint(x[1:6:3])\n"),
        "[1, 4]"
    );
}

#[test]
fn str_slice_step_three() {
    assert_eq!(run_print("'abcdefg'[::3]"), "adg");
}

#[test]
fn list_slice_negative_step() {
    assert_eq!(
        run_python_one("x = [0, 1, 2, 3, 4]\nprint(x[4:1:-1])\n"),
        "[4, 3, 2]"
    );
}

#[test]
fn tuple_slice_negative_step() {
    assert_eq!(
        run_python_one("t = (0, 1, 2, 3)\nprint(t[3:0:-1])\n"),
        "(3, 2, 1)"
    );
}

#[test]
fn list_slice_open_negative_start() {
    assert_eq!(
        run_python_one("x = [1, 2, 3, 4]\nprint(x[-2:])\n"),
        "[3, 4]"
    );
}

#[test]
fn str_slice_open_negative_start() {
    assert_eq!(run_print("'vybe'[-2:]"), "be");
}

#[test]
fn list_slice_len_after_slice() {
    assert_eq!(
        run_python_one("x = [1, 2, 3, 4, 5]\nprint(len(x[1:4]))\n"),
        "3"
    );
}

/// A string slice counts CODE POINTS, matching `len` and `s[i]`
/// (`unifiedstringplan.md` Axis 1 — Python's index unit is `scalar`).
/// A UTF-16 slice would cut the astral pair in half and yield mojibake.
#[test]
fn str_slice_counts_code_points() {
    assert_eq!(
        run_python_one("s = \"a\\U0001F600b\"\nprint(s[1:2], s[0:2], s[1:])\n"),
        "\u{1F600} a\u{1F600} \u{1F600}b"
    );
}

/// Slicing an empty list literal yields an empty list — it is not an
/// out-of-range index. The walker used to raise `IndexError` for every
/// subscript of `[]`, without distinguishing a key from a slice.
#[test]
fn empty_list_literal_slice_is_empty() {
    assert_eq!(run_print("[][0:2]"), "[]");
}

/// Slice bounds are ordinary expressions: an attribute read has to be
/// desugared like any other, not passed through untouched.
#[test]
fn slice_bounds_from_attributes() {
    assert_eq!(
        run_python_one(
            "class B:\n    def __init__(self):\n        self.a = 1\n        self.b = 3\nb = B()\nprint([10, 20, 30, 40][b.a:b.b])\n"
        ),
        "[20, 30]"
    );
}
