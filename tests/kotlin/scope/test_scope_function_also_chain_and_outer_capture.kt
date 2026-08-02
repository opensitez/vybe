// vybe-test: kotlin/scope/test_scope_function_also_chain_and_outer_capture
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var marker = "base"
            val values = mutableListOf(1, 2, 3)
                .also {
                    it.add(4)
                }
                .also {
                    marker = "after"
                }

            __check((values.joinToString(",")).toString(), "1,2,3,4")
            __check((marker).toString(), "after")
        }
