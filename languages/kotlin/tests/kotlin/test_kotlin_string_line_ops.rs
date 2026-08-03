kotlin_run_cases! {
    test_lines_breaks_and_counts => (r##"
        fun main() {
            val value = "a\nb\nc"
            val lines = value.lines()
            println(lines.size)
            println(lines[0])
            println(lines[1])
            println(lines[2])
        }
    "##, &[
        "3",
        "a",
        "b",
        "c",
    ]),
    test_lines_without_trailing => (r##"
        fun main() {
            val value = "a\n"
            val lines = value.lines()
            println(lines.size)
            println(lines[0])
            println(lines[1])
        }
    "##, &[
        "2",
        "a",
        "",
    ]),
    test_line_sequence_counts => (r##"
        fun main() {
            val value = "x\ny\nz\n"
            val count = value.lineSequence().count()
            val tail = value.lineSequence().last()
            println(count)
            println(tail)
        }
    "##, &[
        "4",
        "",
    ]),
    test_substring_before_after => (r##"
        fun main() {
            println("abc:def".substringBefore(":"))
            println("abc:def".substringAfter(":"))
            println("abc".substringBefore("x", "fallback"))
            println("abc".substringAfter("x", "fallback"))
        }
    "##, &[
        "abc",
        "def",
        "fallback",
        "fallback",
    ]),
    test_substring_before_after_last => (r##"
        fun main() {
            println("a:b:c".substringBeforeLast(":"))
            println("a:b:c".substringAfterLast(":"))
            println("a:b:c".substringBeforeLast("x", "fallback"))
            println("a:b:c".substringAfterLast("x", "fallback"))
        }
    "##, &[
        "a:b",
        "c",
        "fallback",
        "fallback",
    ]),
    test_remove_prefix_suffix => (r##"
        fun main() {
            println("prefix:value".removePrefix("prefix:"))
            println("prefix:value".removePrefix("x"))
            println("value/suffix".removeSuffix("/suffix"))
            println("value/suffix".removeSuffix("x"))
        }
    "##, &[
        "value",
        "prefix:value",
        "value",
        "value/suffix",
    ]),
    test_trim_prefix_suffix => (r##"
        fun main() {
            println("   padded   ".trim())
            println("...x".trimStart('.'))
            println("...x".trimEnd('.'))
            println("   x".trimStart())
        }
    "##, &[
        "padded",
        "x",
        "...x",
        "x",
    ]),
    test_trim_margin_and_indent => (r##"
        fun main() {
            val text = """
                |a
                |b
                |c
            """.trimMargin()
            println(text)
            val raw = """\n    one\n    two\n""".trimIndent()
            println(raw.startsWith("one"))
            println(raw.lines().size)
        }
    "##, &[
        "a\nb\nc",
        "true",
        "2",
    ]),
    test_string_sub_sequence => (r##"
        fun main() {
            val s = "abcdef"
            println(s.subSequence(1, 4))
            println(s.substring(2, 4))
            println(s.take(2))
            println(s.drop(2))
        }
    "##, &[
        "bcd",
        "cd",
        "ab",
        "cdef",
    ]),
    test_string_replace_and_split_ops => (r##"
        fun main() {
            println("aa-bb-cc".replaceFirst("-", "/"))
            println("aa-bb-cc".replace("-", "/"))
            val out = "a,b,c".split(",")
            println(out.size)
            println(out[1])
            println("a,b,c".split(",", limit = 2).size)
        }
    "##, &[
        "aa/bb-cc",
        "aa/bb/cc",
        "3",
        "b",
        "2",
    ]) }
