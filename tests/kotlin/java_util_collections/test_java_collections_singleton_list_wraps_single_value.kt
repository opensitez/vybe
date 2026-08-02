// vybe-test: kotlin/java_util_collections/test_java_collections_singleton_list_wraps_single_value
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Collections.singletonList("only")
            __check((values.size).toString(), "1")
            __check((values[0]).toString(), "only")
        }
