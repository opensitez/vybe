// vybe-test: kotlin/exceptions/test_require_false_path_with_message
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
                __check(("require").toString(), "require")
            }
            __check(("done").toString(), "done")
        }
