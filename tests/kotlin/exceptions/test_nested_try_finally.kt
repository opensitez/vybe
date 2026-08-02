// vybe-test: kotlin/exceptions/test_nested_try_finally
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
                    throw Exception("inner")
                } catch (e: Exception) {
                    __check(("inner catch").toString(), "inner catch")
                } finally {
                    __check(("inner finally").toString(), "inner finally")
                }
            } finally {
                __check(("outer finally").toString(), "outer finally")
            }
        }
