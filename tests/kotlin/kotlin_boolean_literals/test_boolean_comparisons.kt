// vybe-test: kotlin/kotlin_boolean_literals/test_boolean_comparisons
// origin: languages/kotlin/tests/kotlin/test_kotlin_boolean_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1 == 1).toString(), "true")
            __check((1 != 2).toString(), "true")
            __check((2 > 1).toString(), "true")
            __check((3 <= 3).toString(), "true")
        }
