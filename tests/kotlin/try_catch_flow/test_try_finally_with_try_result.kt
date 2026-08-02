// vybe-test: kotlin/try_catch_flow/test_try_finally_with_try_result
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = try {
                100
            } finally {
                __check(("f").toString(), "f")
            }
            __check((x).toString(), "100")
        }
