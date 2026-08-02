// vybe-test: kotlin/extension_functions/test_extension_function_for_boolean_returns_numeric_projection
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun Boolean.intValue(): Int = if (this) 1 else 0

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true.intValue()).toString(), "1")
            __check((false.intValue()).toString(), "0")
        }
