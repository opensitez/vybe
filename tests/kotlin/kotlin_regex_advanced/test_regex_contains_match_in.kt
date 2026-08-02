// vybe-test: kotlin/kotlin_regex_advanced/test_regex_contains_match_in
// origin: languages/kotlin/tests/kotlin/test_kotlin_regex_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val regex = Regex("foo")
            __check((regex.containsMatchIn("bar foo baz")).toString(), "true")
            __check((regex.matches("foo")).toString(), "true")
            __check((regex.matches("bar foo baz")).toString(), "false")
        }
