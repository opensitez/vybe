// vybe-test: kotlin/scoping_functions/test_let_can_be_used_for_conditional_mapping
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 8
            val result = if (value > 5) {
                value.let { it * 2 }
            } else {
                0
            }
            __check((result).toString(), "16")
        }
