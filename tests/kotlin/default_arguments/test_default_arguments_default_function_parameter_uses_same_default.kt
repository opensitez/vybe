// vybe-test: kotlin/default_arguments/test_default_arguments_default_function_parameter_uses_same_default
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun fallback(v: String = "z"): String = v
        fun wrapper(label: String, value: String = fallback()): String = label + value
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((wrapper("a")).toString(), "az")
            __check((wrapper("a", "b")).toString(), "ab")
        }
