// vybe-test: kotlin/scoping_functions/test_run_block_can_capture_outer_variables
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var total = 1
            val value = run {
                total += 2
                total * 2
            }
            __check((total).toString(), "3")
            __check((value).toString(), "6")
        }
