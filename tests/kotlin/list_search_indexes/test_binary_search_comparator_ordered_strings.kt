// vybe-test: kotlin/list_search_indexes/test_binary_search_comparator_ordered_strings
// origin: languages/kotlin/tests/kotlin/test_list_search_indexes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("aa", "bb", "cc")
            val byLength = compareBy<String> { it.length }
            __check((values.binarySearch("zz", comparator = byLength)).toString(), "-4")
            __check((values.binarySearch("b", comparator = byLength)).toString(), "0")
        }
