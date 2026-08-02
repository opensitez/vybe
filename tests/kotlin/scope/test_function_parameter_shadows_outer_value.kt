// vybe-test: kotlin/scope/test_function_parameter_shadows_outer_value
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun compute(value: Int): Int {
            return value * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 5
            __check((compute(value)).toString(), "10")
            __check((value).toString(), "5")
        }
