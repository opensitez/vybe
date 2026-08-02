// vybe-test: kotlin/boolean_logic/test_boolean_and_comparison_chains
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 3
            val b = 4
            __check((a < b && a == 3).toString(), "true")
            __check((a > b || b == 4).toString(), "true")
            __check((a <= b && b % 2 == 0).toString(), "false")
            __check((a in 1..b).toString(), "true")
        }
