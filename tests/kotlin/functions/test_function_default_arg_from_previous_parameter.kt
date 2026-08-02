// vybe-test: kotlin/functions/test_function_default_arg_from_previous_parameter
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun tag(base: String, suffix: String = base.uppercase()): String {
            return base + ":" + suffix
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((tag("kotlin")).toString(), "kotlin:KOTLIN")
            __check((tag("kotlin", "custom")).toString(), "kotlin:custom")
        }
