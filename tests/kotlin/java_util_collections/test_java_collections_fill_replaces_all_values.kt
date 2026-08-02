// vybe-test: kotlin/java_util_collections/test_java_collections_fill_replaces_all_values
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 3, 4))
            java.util.Collections.fill(values, 9)
            __check((values).toString(), "[9, 9, 9, 9]")
        }
