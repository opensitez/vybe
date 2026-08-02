// vybe-test: kotlin/kotlin_numeric_literal_formats/test_hex_and_binary_literals
// origin: languages/kotlin/tests/kotlin/test_kotlin_numeric_literal_formats.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val hex = 0x10
            val bin = 0b10
            val oct = 8
            __check((hex).toString(), "16")
            __check((bin).toString(), "2")
            __check((oct).toString(), "8")
        }
