// vybe-test: kotlin/extension_properties/test_extension_property_string_is_numeric
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String.isNumeric: Boolean get() = this.all { ch -> ch in '0'..'9' }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("123".isNumeric).toString(), "true")
            __check(("12a3".isNumeric).toString(), "false")
        }
