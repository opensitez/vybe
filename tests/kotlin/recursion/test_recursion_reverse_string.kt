// vybe-test: kotlin/recursion/test_recursion_reverse_string
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun rev(s: String): String = if (s.isEmpty()) "" else rev(s.substring(1)) + s[0]
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((rev("abc")).toString(), "cba")
        }
