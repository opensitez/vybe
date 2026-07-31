kotlin_run_cases! {
    test_starts_with_basic => (r#"
        fun main() {
            val s = "kotlin"
            println(s.startsWith("kot").toString())
            println(s.startsWith("lin").toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_starts_with_ignorecase => (r#"
        fun main() {
            val s = "Kotlin"
            println(s.startsWith("k", true).toString())
            println(s.startsWith("ko", true).toString())
        }
    "#, vec![String::from("true"), String::from("true")]),
    test_ends_with_basic => (r#"
        fun main() {
            val s = "language"
            println(s.endsWith("age").toString())
            println(s.endsWith("lang").toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_ends_with_case => (r#"
        fun main() {
            val s = "Hello"
            println(s.endsWith("O", true).toString())
        }
    "#, vec![String::from("true")]),
    test_starts_with_empty => (r#"
        fun main() {
            val s = "abc"
            println(s.startsWith("").toString())
            println("".startsWith(s).toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_starts_with_char => (r#"
        fun main() {
            val s = "abc"
            println(s.startsWith("a").toString())
            println(s.endsWith("c").toString())
        }
    "#, vec![String::from("true"), String::from("true")]),
    test_starts_with_offset => (r#"
        fun main() {
            val s = "prefix"
            println(s.startsWith("fix", startIndex = 3).toString())
        }
    "#, vec![String::from("true")]),
    test_ends_with_offset => (r#"
        fun main() {
            val s = "abcXYZ"
            val sub = "XYZ"
            println(s.endsWith(sub).toString())
        }
    "#, vec![String::from("true")]),
    test_starts_end_combo => (r#"
        fun main() {
            val s = "kotlin-language"
            println(if (s.startsWith("kot") && s.endsWith("age")) "both" else "no")
        }
    "#, vec![String::from("both")]),
    test_starts_with_digit => (r#"
        fun main() {
            val s = "123abc"
            println(s.startsWith("1").toString())
            println(s.endsWith("abc").toString())
        }
    "#, vec![String::from("true"), String::from("true")]),
    test_prefix_suffix_empty => (r#"
        fun main() {
            val s = ""
            println(s.startsWith("x").toString())
            println(s.endsWith("").toString())
        }
    "#, vec![String::from("false"), String::from("true")]),
    test_starts_with_unicode => (r#"
        fun main() {
            val s = "Ωmega"
            println(s.startsWith("Ω").toString())
            println(s.endsWith("a").toString())
        }
    "#, vec![String::from("true"), String::from("true")]),
}
