// vybe-test: kotlin/extension_properties/test_extension_property_list_is_empty
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val <T> List<T>.isNoItems: Boolean get() = this.isEmpty()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((listOf<Int>().isNoItems).toString(), "true")
            __check((listOf(1, 2).isNoItems).toString(), "false")
        }
