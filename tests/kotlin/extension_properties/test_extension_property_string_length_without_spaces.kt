// vybe-test: kotlin/extension_properties/test_extension_property_string_length_without_spaces
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String.trimmedLength: Int get() = this.trim().length
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("  a b  ".trimmedLength).toString(), "3")
        }
