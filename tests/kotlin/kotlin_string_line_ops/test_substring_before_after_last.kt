// vybe-test: kotlin/kotlin_string_line_ops/test_substring_before_after_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a:b:c".substringBeforeLast(":")).toString(), "a:b")
            __check(("a:b:c".substringAfterLast(":")).toString(), "c")
            __check(("a:b:c".substringBeforeLast("x", "fallback")).toString(), "fallback")
            __check(("a:b:c".substringAfterLast("x", "fallback")).toString(), "fallback")
        }
