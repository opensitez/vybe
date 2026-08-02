// vybe-test: kotlin/named_arguments/test_named_arguments_boolean_with_shadowed_name
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun emit(enabled: Boolean = false, label: String = "off"): String {
            return if (enabled) "on:" + label else "off:" + label
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val enabled = true
            __check((emit(enabled = enabled, label = "v")).toString(), "on:v")
        }
