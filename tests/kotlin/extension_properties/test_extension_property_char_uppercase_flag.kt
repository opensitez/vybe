// vybe-test: kotlin/extension_properties/test_extension_property_char_uppercase_flag
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Char.isUpper: Boolean get() = this in 'A'..'Z'
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(('R'.isUpper).toString(), "true")
            __check(('t'.isUpper).toString(), "false")
        }
