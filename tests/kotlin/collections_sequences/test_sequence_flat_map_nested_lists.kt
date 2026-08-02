// vybe-test: kotlin/collections_sequences/test_sequence_flat_map_nested_lists
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(listOf(1, 2), listOf(3, 4)).asSequence().flatMap { it.asSequence() }
            __check((seq.sum()).toString(), "10")
            __check((seq.toList().joinToString(",")).toString(), "1,2,3,4")
        }
