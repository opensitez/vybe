// vybe-test: kotlin/kotlin_string_number_parsing/test_parse_byte_bounds
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_number_parsing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("127".toByteOrNull()).toString(), "127")
            __check(("128".toByteOrNull()).toString(), "null")
            __check(("-128".toByte()).toString(), "-128")
            __check(("-129".toByteOrNull()).toString(), "null")
        }
