// vybe-test: kotlin/boolean_logic/test_boolean_xor_behavior
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true xor true).toString(), "false")
            __check((true xor false).toString(), "true")
            __check((false xor true).toString(), "true")
            __check((false xor false).toString(), "false")
        }
