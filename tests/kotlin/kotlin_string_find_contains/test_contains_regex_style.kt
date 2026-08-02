// vybe-test: kotlin/kotlin_string_find_contains/test_contains_regex_style
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "abc123def"
            __check((s.indexOf("123").toString()).toString(), "3")
            __check((s.contains("123").toString()).toString(), "true")
        }
