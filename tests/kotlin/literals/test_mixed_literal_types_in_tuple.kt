// vybe-test: kotlin/literals/test_mixed_literal_types_in_tuple
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2.0, true, 'x')
            __check((values.size).toString(), "4")
            __check((values[0]).toString(), "1")
            __check((values[1]).toString(), "2.0")
            __check((values[2]).toString(), "true")
            __check((values[3]).toString(), "x")
        }
