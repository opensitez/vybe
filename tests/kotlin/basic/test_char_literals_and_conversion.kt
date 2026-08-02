// vybe-test: kotlin/basic/test_char_literals_and_conversion
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first: Char = 'A'
            val second = 'b'
            __check((first).toString(), "A")
            __check((second).toString(), "b")
            __check((first.code + 1).toString(), "66")
            __check((second.code).toString(), "98")
        }
