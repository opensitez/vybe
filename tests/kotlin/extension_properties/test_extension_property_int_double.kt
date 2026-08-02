// vybe-test: kotlin/extension_properties/test_extension_property_int_double
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Int.doubled: Int get() = this * 2
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((7.doubled).toString(), "14")
        }
