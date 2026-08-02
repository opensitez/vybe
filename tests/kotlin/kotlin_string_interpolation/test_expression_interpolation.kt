// vybe-test: kotlin/kotlin_string_interpolation/test_expression_interpolation
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 3
            __check(("sum " + "${value + 1}").toString(), "sum 4")
        }
