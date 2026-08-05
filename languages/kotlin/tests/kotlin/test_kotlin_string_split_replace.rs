kotlin_run_cases! {
    test_split_single => (r#"
        fun main() {
            val s = "a,b,c"
            val parts = s.split(",")
            println(parts.size)
            println(parts[1])
        }
    "#, vec![String::from("3"), String::from("b")]),
    test_split_limit => (r#"
        fun main() {
            val s = "a,b,c,d"
            val parts = s.split(",", limit = 2)
            println(parts.size)
            println(parts[1])
        }
    "#, vec![String::from("2"), String::from("b,c,d")]),
    test_split_whitespace => (r#"
        fun main() {
            val s = "one two  three"
            val parts = s.split(" ")
            println(parts.size)
            println(parts[0])
        }
    "#, vec![String::from("4"), String::from("one")]),
    test_split_multiple_delim => (r#"
        fun main() {
            val s = "a|b|c"
            val parts = s.split("|")
            println(parts[2])
        }
    "#, vec![String::from("c")]),
    test_split_empty_parts => (r#"
        fun main() {
            val s = "a,,b"
            val parts = s.split(",", limit = 0)
            println(parts.size)
            println(parts[1])
        }
    "#, vec![String::from("3"), String::from("")]),
    test_split_regex_like => (r#"
        fun main() {
            val s = "1;2;3"
            val parts = s.split(";")
            println(parts[0] + parts[2])
        }
    "#, vec![String::from("13")]),
    test_replace_basic => (r#"
        fun main() {
            val s = "abc"
            println(s.replace("a", "x"))
        }
    "#, vec![String::from("xbc")]),
    test_replace_first => (r#"
        fun main() {
            val s = "banana"
            println(s.replaceFirst("ana", "x"))
        }
    "#, vec![String::from("bxna")]),
    test_replace_two => (r#"
        fun main() {
            val s = "aa bb aa"
            println(s.replace("aa", "x"))
        }
    "#, vec![String::from("x bb x")]),
    test_replace_without_match => (r#"
        fun main() {
            val s = "abc"
            println(s.replace("z", "q"))
        }
    "#, vec![String::from("abc")]),
    test_replace_range => (r#"
        fun main() {
            val s = "abcdef"
            val out = StringBuilder(s).replace(1, 3, "ZZ")
            println(out.toString())
        }
    "#, vec![String::from("aZZdef")]),
    test_split_plus_join => (r#"
        fun main() {
            val s = "a,b,c"
            val joined = s.split(",").joinToString("-")
            println(joined)
        }
    "#, vec![String::from("a-b-c")]),
}
