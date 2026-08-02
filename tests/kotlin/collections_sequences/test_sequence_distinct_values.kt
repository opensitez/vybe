// vybe-test: kotlin/collections_sequences/test_sequence_distinct_values
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 1, 2, 2, 3).asSequence().distinct()
            __check((seq.toList().joinToString(",")).toString(), "1,2,3")
        }
