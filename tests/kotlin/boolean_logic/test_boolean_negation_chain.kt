// vybe-test: kotlin/boolean_logic/test_boolean_negation_chain
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = true
            val b = false
            __check((!a).toString(), "false")
            __check((!!a).toString(), "true")
            __check((!!!b).toString(), "true")
            __check((!(!a && b)).toString(), "true")
        }
