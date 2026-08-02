// vybe-test: kotlin/exceptions/test_require_helper
// origin: languages/kotlin/tests/kotlin/test_exceptions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                require(false)
            } catch (e: Exception) {
                __check(("require failed").toString(), "require failed")
            }
        }
