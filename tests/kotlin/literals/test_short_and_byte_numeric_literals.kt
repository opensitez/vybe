// vybe-test: kotlin/literals/test_short_and_byte_numeric_literals
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val small: Short = 12
            val tiny: Byte = 7
            val unsigned: Int = 1
            __check((small).toString(), "12")
            __check((tiny).toString(), "7")
            __check((unsigned).toString(), "1")
            __check((small + tiny + unsigned).toString(), "20")
        }
