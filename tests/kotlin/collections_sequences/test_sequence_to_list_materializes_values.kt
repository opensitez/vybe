// vybe-test: kotlin/collections_sequences/test_sequence_to_list_materializes_values
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3).map { it * 10 }
            val out = seq.toList()
            __check((out.joinToString(",")).toString(), "10,20,30")
        }
