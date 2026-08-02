// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_with_callable_reference_bound
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

fun <T : Number> describe(value: T): String {
            return value.toString() + value.toInt().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val fn: (Int) -> String = ::describe
            __check((fn(3)).toString(), "33")
        }
