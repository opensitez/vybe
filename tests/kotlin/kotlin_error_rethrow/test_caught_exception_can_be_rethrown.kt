// vybe-test: kotlin/kotlin_error_rethrow/test_caught_exception_can_be_rethrown
// origin: languages/kotlin/tests/kotlin/test_kotlin_error_rethrow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                try {
                    throw Exception("inner")
                } catch (e: Exception) {
                    __check(("inner").toString(), "inner")
                    throw e
                }
            } catch (e: Exception) {
                __check(("outer").toString(), "outer")
            }
        }
