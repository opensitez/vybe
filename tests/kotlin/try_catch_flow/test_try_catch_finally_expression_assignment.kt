// vybe-test: kotlin/try_catch_flow/test_try_catch_finally_expression_assignment
// origin: languages/kotlin/tests/kotlin/test_try_catch_flow.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = try {
                throw Exception("x")
            } catch (e: Exception) {
                8
            } finally {
                __check(("cleanup").toString(), "cleanup")
            }
            __check((value).toString(), "8")
        }
