// vybe-test: kotlin/try_catch_flow/test_try_nested_catch
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

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
