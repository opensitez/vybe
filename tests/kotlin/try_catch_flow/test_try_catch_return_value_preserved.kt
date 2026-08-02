// vybe-test: kotlin/try_catch_flow/test_try_catch_return_value_preserved
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun safeDivide(a: Int, b: Int): Int {
            return try {
                a / b
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
            __check((safeDivide(10, 2)).toString(), "5")
            __check((safeDivide(10, 0)).toString(), "0")
        }
