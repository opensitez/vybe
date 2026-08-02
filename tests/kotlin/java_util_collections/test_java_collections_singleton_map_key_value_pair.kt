// vybe-test: kotlin/java_util_collections/test_java_collections_singleton_map_key_value_pair
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val entry = java.util.Collections.singletonMap("a", 1)
            __check((entry.size).toString(), "1")
            __check((entry["a"]).toString(), "1")
            __check((entry["missing"] ?: "none").toString(), "none")
        }
