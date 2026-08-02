// vybe-test: kotlin/extension_properties/test_extension_property_int_triple
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Int.tripled: Int get() = this * 3
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((4.tripled).toString(), "12")
        }
