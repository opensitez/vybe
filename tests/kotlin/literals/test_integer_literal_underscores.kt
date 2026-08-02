// vybe-test: kotlin/literals/test_integer_literal_underscores
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((1_000).toString(), "1000")
            __check((10_000_000).toString(), "10000000")
            __check((1_2_3_4).toString(), "1234")
        }
