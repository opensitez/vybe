// vybe-test: kotlin/collections_iterables/test_list_min_by_and_max_by
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("pear", "apple", "banana")
            __check((words.minByOrNull { it.length } ?: "").toString(), "pear")
            __check((words.maxByOrNull { it.length } ?: "").toString(), "banana")
        }
