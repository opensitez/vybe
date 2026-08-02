// vybe-test: kotlin/scope/test_scope_try_catch_variable_not_visible_after_block
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val status = "start"
            val result = try {
                throw Exception("boom")
            } catch (failure: Exception) {
                val status = "caught"
                status
            } finally {
                __check((status).toString(), "start")
            }
            __check((result).toString(), "caught")
        }
