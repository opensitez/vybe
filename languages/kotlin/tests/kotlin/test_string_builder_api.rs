use crate::helpers::run_prints;

#[test]
fn test_string_builder_append_values() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            out.append("a").append(1).append('b')
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["a1b"]);
}

#[test]
fn test_string_builder_insert() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("ab")
            out.insert(1, "X")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["aXb"]);
}

#[test]
fn test_string_builder_delete() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abcd")
            out.delete(1, 3)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["ad"]);
}

#[test]
fn test_string_builder_delete_at() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abc")
            out.deleteAt(1)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["ac"]);
}

#[test]
fn test_string_builder_set_char_at() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abc")
            out.setCharAt(1, 'Z')
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["aZc"]);
}

#[test]
fn test_string_builder_append_line() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            out.appendLine("a").appendLine("b")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["a\nb\n"]);
}

#[test]
fn test_string_builder_reverse_and_length() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("kotlin")
            println(out.length)
            out.reverse()
            println(out.toString())
            println(out.length)
        }
    "#,
    );
    assert_eq!(out, &["6", "niltok", "6"]);
}

#[test]
fn test_string_builder_sub_sequence() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abcdef")
            println(out.substring(1, 4))
            println(out.subSequence(2, 5))
        }
    "#,
    );
    assert_eq!(out, &["bcd", "cde"]);
}

#[test]
fn test_string_builder_capacity_growth() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder(4)
            out.append("abcd")
            println(out.capacity() >= 4)
            out.append("ef")
            println(out.capacity() >= 6)
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_string_builder_set_length_truncate() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("hello")
            out.setLength(2)
            println(out.toString())
            println(out.length)
        }
    "#,
    );
    assert_eq!(out, &["he", "2"]);
}

#[test]
fn test_string_builder_clear_via_set_length_zero() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("clear")
            out.setLength(0)
            println(out.isEmpty())
            println(out.length)
        }
    "#,
    );
    assert_eq!(out, &["true", "0"]);
}

#[test]
fn test_string_builder_replace_range() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("a-b-c")
            out.replace(1, 4, "B")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["aBc"]);
}

#[test]
fn test_string_builder_append_format() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            val value = 7
            out.append("value=").append(value).append(",")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["value=7,"]);
}

#[test]
fn test_string_builder_append_code_point_style() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            out.appendCodePoint(97)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["a"]);
}

#[test]
fn test_string_builder_indices_navigation() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abc")
            println(out[0])
            println(out[1])
            println(out[2])
        }
    "#,
    );
    assert_eq!(out, &["a", "b", "c"]);
}

#[test]
fn test_string_builder_append_range_like_primitive() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            val chars = charArrayOf('x', 'y', 'z')
            out.append(chars)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["xyz"]);
}

#[test]
fn test_string_builder_appendln_alias() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            out.append("x").appendLine()
            out.append("y")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["x\ny"]);
}

#[test]
fn test_string_builder_repeat_append_ints() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            repeat(3) { out.append('x') }
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["xxx"]);
}

#[test]
fn test_string_builder_to_string_stability() {
    let out = run_prints(
        r#"
        fun main() {
            val base = StringBuilder("abc")
            val text = base.toString()
            base.append("d")
            println(text)
            println(base.toString())
        }
    "#,
    );
    assert_eq!(out, &["abc", "abcd"]);
}

#[test]
fn test_string_builder_build_from_values() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            val a = listOf("x", "y", "z")
            for (item in a) {
                out.append(item)
            }
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["xyz"]);
}

#[test]
fn test_string_builder_trimmed_content() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("  abc  ")
            println(out.toString().trim())
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["abc", "  abc  "]);
}

#[test]
fn test_string_builder_append_other_builders() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            val inner = StringBuilder("inner")
            out.append(inner)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["inner"]);
}

#[test]
fn test_string_builder_last_index() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("tool")
            println(out.lastIndex)
            println(out[out.lastIndex])
        }
    "#,
    );
    assert_eq!(out, &["3", "l"]);
}

#[test]
fn test_string_builder_length_and_hashcode() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("hash")
            println(out.length)
            println(out.toString().hashCode() == out.hashCode())
        }
    "#,
    );
    assert_eq!(out, &["4", "false"]);
}

#[test]
fn test_string_builder_clear_then_append() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder("abc")
            out.setLength(0)
            out.append("x")
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["x"]);
}

#[test]
fn test_string_builder_join_chars_from_list() {
    let out = run_prints(
        r#"
        fun main() {
            val out = StringBuilder()
            val chars = listOf('a', 'b', 'c')
            for (c in chars) out.append(c)
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}
