// vybe-test: kotlin/kotlin_string_line_ops/test_substring_before_after
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("abc:def".substringBefore(":")).toString(), "abc")
            __check(("abc:def".substringAfter(":")).toString(), "def")
            __check(("abc".substringBefore("x", "fallback")).toString(), "fallback")
            __check(("abc".substringAfter("x", "fallback")).toString(), "fallback")
        }
