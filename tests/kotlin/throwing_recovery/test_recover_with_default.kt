// vybe-test: kotlin/throwing_recovery/test_recover_with_default
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun fallback(x: String): Int {
            try {
                return x.toInt()
            } catch (e: NumberFormatException) {
                return 0
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((fallback("12")).toString(), "12")
            __check((fallback("xx")).toString(), "0")
        }
