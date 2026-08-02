// vybe-test: kotlin/extension_properties/test_extension_property_byte_is_zero
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Byte.isZero: Boolean get() = this == 0.toByte()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0.toByte().isZero).toString(), "true")
            __check((2.toByte().isZero).toString(), "false")
        }
