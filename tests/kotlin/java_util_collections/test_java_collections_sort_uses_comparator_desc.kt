// vybe-test: kotlin/java_util_collections/test_java_collections_sort_uses_comparator_desc
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = java.util.ArrayList<String>(listOf("bbb", "c", "aa"))
            java.util.Collections.sort(values, java.util.Collections.reverseOrder<String>())
            __check((values).toString(), "[c, bbb, aa]")
        }
