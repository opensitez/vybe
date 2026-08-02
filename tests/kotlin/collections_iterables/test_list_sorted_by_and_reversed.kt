// vybe-test: kotlin/collections_iterables/test_list_sorted_by_and_reversed
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val letters = listOf("pear", "apple", "kiwi")
            __check((letters.sorted().joinToString(",")).toString(), "apple,kiwi,pear")
            __check((letters.reversed().joinToString(",")).toString(), "kiwi,apple,pear")
        }
