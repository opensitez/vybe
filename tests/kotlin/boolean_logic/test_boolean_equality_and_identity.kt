// vybe-test: kotlin/boolean_logic/test_boolean_equality_and_identity
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true == true).toString(), "true")
            __check((true == false).toString(), "false")
            __check((false == false).toString(), "true")
            val a: Boolean = false
            val b = a
            __check((a === b).toString(), "true")
        }
