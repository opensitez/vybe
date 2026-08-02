// vybe-test: kotlin/default_arguments/test_default_arguments_defaulted_boolean_chain
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun flags(a: Boolean = true, b: Boolean = false): String = if (a && !b) "on" else "off"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((flags()).toString(), "on")
            __check((flags(a = false)).toString(), "off")
            __check((flags(b = true)).toString(), "off")
        }
