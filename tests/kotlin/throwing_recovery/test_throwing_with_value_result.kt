// vybe-test: kotlin/throwing_recovery/test_throwing_with_value_result
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun recover(v: Int): Int {
            return try {
                if (v == 0) throw Exception("x")
                100 / v
            } catch (e: Exception) {
                0
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((recover(4)).toString(), "25")
            __check((recover(0)).toString(), "0")
        }
