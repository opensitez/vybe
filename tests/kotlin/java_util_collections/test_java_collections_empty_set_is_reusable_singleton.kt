// vybe-test: kotlin/java_util_collections/test_java_collections_empty_set_is_reusable_singleton
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Collections.emptySet<String>()
            __check((values.isEmpty()).toString(), "true")
            __check((values.contains("a")).toString(), "false")
        }
