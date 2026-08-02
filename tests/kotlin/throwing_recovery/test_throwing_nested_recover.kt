// vybe-test: kotlin/throwing_recovery/test_throwing_nested_recover
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun parseIntOrThrow(s: String): Int {
            if (s == "x") throw NumberFormatException("bad")
            return 1
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                parseIntOrThrow("x")
            } catch (e: NumberFormatException) {
                __check(("bad").toString(), "bad")
            }
        }
