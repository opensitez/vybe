// vybe-test: kotlin/literals/test_float_suffix_literal
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tiny: Float = 1.25f
            val rounded = tiny + 0.75f
            __check((tiny).toString(), "1.25")
            __check((rounded).toString(), "2.0")
            __check((rounded.toString()).toString(), "2.0")
        }
