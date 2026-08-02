// vybe-test: kotlin/extension_properties/test_extension_property_int_is_even
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Int.isEven: Boolean get() = this % 2 == 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((4.isEven).toString(), "true")
            __check((5.isEven).toString(), "false")
        }
