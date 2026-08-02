// vybe-test: kotlin/extension_properties/test_extension_property_char_digit_flag
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Char.isDigitAscii: Boolean get() = this in '0'..'9'
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('3'.isDigitAscii).toString(), "true")
            __check(('x'.isDigitAscii).toString(), "false")
        }
