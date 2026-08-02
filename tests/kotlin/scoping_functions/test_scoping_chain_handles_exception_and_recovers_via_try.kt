// vybe-test: kotlin/scoping_functions/test_scoping_chain_handles_exception_and_recovers_via_try
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = try {
                5.let {
                    if (it < 0) throw RuntimeException("x")
                    it * 2
                }
            } catch (error: RuntimeException) {
                -1
            }
            __check((result).toString(), "10")
        }
