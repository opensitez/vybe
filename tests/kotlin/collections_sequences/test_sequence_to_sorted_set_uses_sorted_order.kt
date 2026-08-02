// vybe-test: kotlin/collections_sequences/test_sequence_to_sorted_set_uses_sorted_order
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(3, 1, 4, 1, 5, 2).toSortedSet()
            __check((seq.joinToString(",")).toString(), "1,2,3,4,5")
        }
