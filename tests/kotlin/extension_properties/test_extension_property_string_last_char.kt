// vybe-test: kotlin/extension_properties/test_extension_property_string_last_char
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String.lastChar: Char get() = this[this.length - 1]
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("rust".lastChar).toString(), "t")
        }
