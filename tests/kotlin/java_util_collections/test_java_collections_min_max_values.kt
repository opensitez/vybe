// vybe-test: kotlin/java_util_collections/test_java_collections_min_max_values
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(8, 1, 4, 3))
            __check((java.util.Collections.min(values)).toString(), "1")
            __check((java.util.Collections.max(values)).toString(), "8")
        }
