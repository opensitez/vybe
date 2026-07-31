kotlin_run_cases! {
    test_unicode_character => (r#"
        fun main() {
            val g = '\u0047'
            val omega = '\u03A9'
            println(g)
            println(omega)
        }
    "#, vec!["G", "Ω"]),
    test_escaped_quotes => (r#"
        fun main() {
            val quote = '\''
            val backslash = '\\'
            println(quote)
            println(backslash)
        }
    "#, vec!["'", "\\"]),
    test_digit_and_letter_chars => (r#"
        fun main() {
            val c1 = '1'
            val c2 = 'A'
            println(c1.toString() + c2.toString())
        }
    "#, vec!["1A"]),
}
