// vybe-test: kotlin/try_catch_flow/test_try_finally_without_exception_result
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = try {
                3 * 4
            } finally {
                __check(("cleanup").toString(), "cleanup")
            }
            __check((x).toString(), "12")
        }
