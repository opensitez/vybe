// vybe-test: kotlin/boolean_logic/test_boolean_negated_predicates_for_classification
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(-2, -1, 0, 1, 2)
            val positives = values.filter { it > 0 }
            val nonPositive = values.filter { !(it > 0) }
            __check((positives.joinToString(",")).toString(), "1,2")
            __check((nonPositive.joinToString(",")).toString(), "-2,-1,0")
            __check((nonPositive.all { ! (it > 0) }).toString(), "true")
        }
