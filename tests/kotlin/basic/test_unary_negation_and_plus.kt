// vybe-test: kotlin/basic/test_unary_negation_and_plus
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val positive = +5
            val negative = -positive
            __check((negative).toString(), "-5")
        }
