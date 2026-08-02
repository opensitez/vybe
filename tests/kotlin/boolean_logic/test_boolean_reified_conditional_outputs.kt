// vybe-test: kotlin/boolean_logic/test_boolean_reified_conditional_outputs
// origin: languages/kotlin/tests/kotlin/test_boolean_logic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun status(flag: Boolean): String {
                return if (!flag) "off" else "on"
            }
            __check((status(false)).toString(), "off")
            __check((status(true)).toString(), "on")
            __check((status(status(false) == "off")).toString(), "on")
        }
