// vybe-test: kotlin/extension_properties/test_extension_property_double_rounded
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Double.roundToEven: Int get() = kotlin.math.roundToInt(this)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((3.4.roundToEven).toString(), "3")
            __check((2.6.roundToEven).toString(), "3")
        }
