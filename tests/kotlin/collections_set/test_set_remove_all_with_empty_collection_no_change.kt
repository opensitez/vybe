// vybe-test: kotlin/collections_set/test_set_remove_all_with_empty_collection_no_change
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.removeAll(emptySet<Int>())).toString(), "false")
            __check((values.size).toString(), "3")
        }
