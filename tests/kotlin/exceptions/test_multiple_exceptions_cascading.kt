// vybe-test: kotlin/exceptions/test_multiple_exceptions_cascading
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                try {
                    throw IllegalArgumentException("bad")
                } catch (e: IllegalArgumentException) {
                    throw Exception("wrapped")
                }
            } catch (e: Exception) {
                __check(("wrapped").toString(), "wrapped")
            }
        }
