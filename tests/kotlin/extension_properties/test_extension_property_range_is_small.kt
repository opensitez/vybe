// vybe-test: kotlin/extension_properties/test_extension_property_range_is_small
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val IntRange.isSmall: Boolean get() = this.count() <= 3
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((1..3).isSmall).toString(), "true")
            __check(((1..5).isSmall).toString(), "false")
        }
