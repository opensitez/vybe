// vybe-test: kotlin/exceptions/test_try_catch_finally_with_return_value
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun calc(): Int {
            try {
                return 5
            } catch (e: Exception) {
                return 0
            } finally {
                __check(("finally").toString(), "finally")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((calc()).toString(), "5")
        }
