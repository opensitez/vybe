// vybe-test: kotlin/java_util_collections/test_java_collections_frequency_counts_occurrences
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.Arrays.asList(1, 2, 1, 3, 1, 2)
            __check((java.util.Collections.frequency(values, 1)).toString(), "3")
            __check((java.util.Collections.frequency(values, 2)).toString(), "2")
            __check((java.util.Collections.frequency(values, 4)).toString(), "0")
        }
