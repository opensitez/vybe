// vybe-test: kotlin/apply_scope_functions/test_run_with_non_trivial_return_type
// origin: languages/kotlin/tests/kotlin/test_apply_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3)
            val score = values.run {
                filter { it % 2 == 1 }.sum()
            }
            __check((score).toString(), "4")
        }
