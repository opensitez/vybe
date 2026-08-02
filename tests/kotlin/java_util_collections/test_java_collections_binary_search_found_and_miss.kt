// vybe-test: kotlin/java_util_collections/test_java_collections_binary_search_found_and_miss
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 5, 8))
            __check((java.util.Collections.binarySearch(values, 3)).toString(), "2")
            __check((java.util.Collections.binarySearch(values, 4)).toString(), "-4")
        }
