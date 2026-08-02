// vybe-test: kotlin/extension_properties/test_extension_property_array_as_text
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val IntArray.totalText: String get() = this.joinToString(",")
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((intArrayOf(1, 2, 3).totalText).toString(), "1,2,3")
        }
