// vybe-test: kotlin/throwing_recovery/test_throwing_and_return_in_catch
// origin: languages/kotlin/tests/kotlin/test_throwing_recovery.rs

fun check(x: Int): Int {
            try {
                if (x == 0) throw Exception("zero")
                return x
            } catch (e: Exception) {
                return 99
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((check(1)).toString(), "1")
            __check((check(0)).toString(), "99")
        }
