// vybe-test: kotlin/local_functions/test_local_function_uses_its_parameter_names_as_types
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun compose(prefix: String, suffix: String): String {
                fun join(value: String): String = prefix + value + suffix
                return join("mid")
            }
            __check((compose("<", ">")).toString(), "<mid>")
        }
