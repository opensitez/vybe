// vybe-test: kotlin/java_util_collections/test_java_collections_replace_all_values
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<Int>(listOf(1, 2, 1, 3))
            java.util.Collections.replaceAll(values, 1, 9)
            __check((values).toString(), "[9, 2, 9, 3]")
        }
