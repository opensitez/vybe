// vybe-test: kotlin/extension_properties/test_extension_property_map_is_present
// origin: languages/kotlin/tests/kotlin/test_extension_properties.rs

val <K, V> Map<K, V>.isPresent: Boolean get() = !isEmpty()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((mapOf<String, Int>().isPresent).toString(), "false")
            __check((mapOf("a" to 1).isPresent).toString(), "true")
        }
