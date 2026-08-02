// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_with_nullable_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun label(value: String?): String {
            return "${'$'}{value ?: "none"}"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((label(null)).toString(), "none")
            __check((label("x")).toString(), "x")
        }
