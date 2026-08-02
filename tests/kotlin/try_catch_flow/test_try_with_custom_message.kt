// vybe-test: kotlin/try_catch_flow/test_try_with_custom_message
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            try {
                throw IllegalArgumentException("bad")
            } catch (e: IllegalArgumentException) {
                __check((e.message).toString(), "bad")
            }
        }
