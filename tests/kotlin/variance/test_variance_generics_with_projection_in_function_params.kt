// vybe-test: kotlin/variance/test_variance_generics_with_projection_in_function_params
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun pick(values: List<out Any>): String {
            return values[0].toString()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pick(listOf(1, 2))).toString(), "1")
            __check((pick(listOf("x", "y"))).toString(), "x")
        }
