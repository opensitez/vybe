// vybe-test: kotlin/java_util_collections/test_java_collections_sort_natural_order
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(4, 1, 3, 2))
            java.util.Collections.sort(values)
            __check((values).toString(), "[1, 2, 3, 4]")
        }
