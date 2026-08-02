// vybe-test: kotlin/java_util_collections/test_java_collections_empty_map_has_zero_size
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Collections.emptyMap<String, Int>()
            __check((values.isEmpty()).toString(), "true")
            __check((values.size).toString(), "0")
        }
