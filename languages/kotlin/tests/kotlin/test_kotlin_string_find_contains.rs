kotlin_run_cases! {
    test_contains_char => (r#"
        fun main() {
            val s = "abcdef"
            println(s.contains("c").toString())
            println(s.contains("z").toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_contains_substring => (r#"
        fun main() {
            val s = "hello world"
            println(s.contains("lo wo").toString())
            println(s.contains("planet").toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_contains_ignorecase => (r#"
        fun main() {
            val s = "Case"
            println(s.contains("c", ignoreCase = true).toString())
            println(s.contains("X", ignoreCase = true).toString())
        }
    "#, vec![String::from("true"), String::from("false")]),
    test_index_of_first => (r#"
        fun main() {
            val s = "banana"
            println(s.indexOf("na").toString())
            println(s.indexOf("na", startIndex = 3).toString())
        }
    "#, vec![String::from("2"), String::from("4")]),
    test_last_index_of => (r#"
        fun main() {
            val s = "banana"
            println(s.lastIndexOf("na").toString())
            println(s.lastIndexOf("x").toString())
        }
    "#, vec![String::from("4"), String::from("-1")]),
    test_index_of_char => (r#"
        fun main() {
            val s = "abcdef"
            println(s.indexOf('d').toString())
            println(s.lastIndexOf('a').toString())
        }
    "#, vec![String::from("3"), String::from("0")]),
    test_find_prefix => (r#"
        fun main() {
            val s = "abcabc"
            println(s.indexOf('a', startIndex = 1).toString())
            println(s.substringAfter("ab").toString())
        }
    "#, vec![String::from("3"), String::from("cabc")]),
    test_substring_range => (r#"
        fun main() {
            val s = "abcdef"
            println(s.substring(1, 3))
            println(s.substring(2))
        }
    "#, vec![String::from("bc"), String::from("cdef")]),
    test_match_indexes => (r#"
        fun main() {
            val s = "x-ay-x"
            println(s.indexOf("-").toString())
            println(s.lastIndexOf("x").toString())
        }
    "#, vec![String::from("1"), String::from("5")]),
    test_contains_none => (r#"
        fun main() {
            val s = ""
            println(s.contains("a").toString())
            println(s.isEmpty().toString())
        }
    "#, vec![String::from("false"), String::from("true")]),
    test_contains_regex_style => (r#"
        fun main() {
            val s = "abc123def"
            println(s.indexOf("123").toString())
            println(s.contains("123").toString())
        }
    "#, vec![String::from("3"), String::from("true")]),
    test_starts_ends_combo => (r#"
        fun main() {
            val s = "prefix_value_suffix"
            val starts = s.startsWith("prefix").toString()
            val ends = s.endsWith("suffix").toString()
            println(starts)
            println(ends)
        }
    "#, vec![String::from("true"), String::from("true")]),
}
