// vybe-test: kotlin/collections_sequences/test_sequence_to_set_and_to_list_stability
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 2, 2, 3).asSequence()
            val asSet = seq.toSet()
            val asList = seq.toList()
            __check((asSet.joinToString(",")).toString(), "1,2,3")
            __check((asList.size).toString(), "4")
        }
