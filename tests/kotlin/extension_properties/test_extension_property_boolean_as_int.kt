// vybe-test: kotlin/extension_properties/test_extension_property_boolean_as_int
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val Boolean.asInt: Int get() = if (this) 1 else 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((true.asInt).toString(), "1")
            __check((false.asInt).toString(), "0")
        }
