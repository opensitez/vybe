// vybe-test: kotlin/scoping_functions/test_run_block_executes_and_returns_value
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val total = run {
                val first = 4
                val second = 6
                first + second
            }
            __check((total).toString(), "10")
        }
