// vybe-test: kotlin/extension_properties/test_extension_property_range_span
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val IntRange.span: Int get() = this.last - this.first
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1..5).span).toString(), "4")
        }
