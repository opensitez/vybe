// vybe-test: kotlin/throwing_recovery/test_throwing_divide_by_zero_recovery
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun safeDivide(a: Int, b: Int): Int {
            try {
                return a / b
            } catch (e: Exception) {
                return -1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((safeDivide(5, 0)).toString(), "-1")
        }
