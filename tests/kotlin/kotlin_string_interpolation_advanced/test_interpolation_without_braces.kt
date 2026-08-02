// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_without_braces
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val count = 4
            __check(("${'$'}count items").toString(), "4 items")
        }
