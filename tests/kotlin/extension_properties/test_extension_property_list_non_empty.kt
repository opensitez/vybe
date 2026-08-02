// vybe-test: kotlin/extension_properties/test_extension_property_list_non_empty
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val List<*>.isNonEmpty: Boolean get() = this.isNotEmpty()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf<Int>().isNonEmpty).toString(), "false")
            __check((listOf(1, 2).isNonEmpty).toString(), "true")
        }
