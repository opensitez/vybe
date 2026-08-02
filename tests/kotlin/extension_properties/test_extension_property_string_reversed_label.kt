// vybe-test: kotlin/extension_properties/test_extension_property_string_reversed_label
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String.reversedLabel: String get() = this.reversed()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("abc".reversedLabel).toString(), "cba")
        }
