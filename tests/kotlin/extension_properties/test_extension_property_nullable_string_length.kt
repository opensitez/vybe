// vybe-test: kotlin/extension_properties/test_extension_property_nullable_string_length
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val String?.orZeroLength: Int get() = this?.length ?: 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x: String? = null
            val y: String? = "kotlin"
            __check((x.orZeroLength).toString(), "0")
            __check((y.orZeroLength).toString(), "6")
        }
