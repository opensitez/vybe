// vybe-test: kotlin/literals/test_integer_literals_hex_and_binary
// origin: languages/kotlin/tests/kotlin/test_literals.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0x1A).toString(), "26")
            __check((0x10 + 1).toString(), "17")
            __check((0b1010).toString(), "10")
            __check((0b1111 + 1).toString(), "16")
        }
