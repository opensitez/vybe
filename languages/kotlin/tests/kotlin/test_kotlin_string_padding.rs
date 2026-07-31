kotlin_run_cases! {
    test_pad_start_default => (r#"
        fun main() {
            val s = "5"
            println(s.padStart(3))
        }
    "#, vec![String::from("  5")]),
    test_pad_start_char => (r#"
        fun main() {
            val s = "7"
            println(s.padStart(3, '0'))
        }
    "#, vec![String::from("007")]),
    test_pad_end_default => (r#"
        fun main() {
            val s = "5"
            println(s.padEnd(3))
        }
    "#, vec![String::from("5  ")]),
    test_pad_end_char => (r#"
        fun main() {
            val s = "7"
            println(s.padEnd(4, '!'))
        }
    "#, vec![String::from("7!!!")]),
    test_pad_start_already_wide => (r#"
        fun main() {
            val s = "abc"
            println(s.padStart(2))
        }
    "#, vec![String::from("abc")]),
    test_pad_end_empty => (r#"
        fun main() {
            val s = ""
            println(s.padStart(2, 'x'))
            println(s.padEnd(2, 'y'))
        }
    "#, vec![String::from("xx"), String::from("yy")]),
    test_repeat_short => (r#"
        fun main() {
            val s = "a"
            println(s.repeat(3))
        }
    "#, vec![String::from("aaa")]),
    test_repeat_zero => (r#"
        fun main() {
            val s = "a"
            println(s.repeat(0))
        }
    "#, vec![String::from("")]),
    test_trim_standard => (r#"
        fun main() {
            val s = "  abc  "
            println(s.trim())
        }
    "#, vec![String::from("abc")]),
    test_trim_start => (r#"
        fun main() {
            val s = "  abc  "
            println(s.trimStart())
        }
    "#, vec![String::from("abc  ")]),
    test_trim_end => (r#"
        fun main() {
            val s = "  abc  "
            println(s.trimEnd())
        }
    "#, vec![String::from("  abc")]),
    test_trim_with_chars => (r#"
        fun main() {
            val s = "--abc--"
            println(s.trim('-'))
        }
    "#, vec![String::from("abc")]),
}
