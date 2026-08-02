// vybe-test: kotlin/collections_sequences/test_sequence_sorted_and_reversed
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("pear", "apple", "kiwi").asSequence()
            __check((source.sorted().joinToString(",")).toString(), "apple,kiwi,pear")
            __check((source.sortedDescending().joinToString(",")).toString(), "pear,kiwi,apple")
        }
