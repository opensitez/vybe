// vybe-test: kotlin/collections_sequences/test_sequence_zip_with_next_returns_adjacent_pairs
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..5).asSequence().zipWithNext { a, b -> a * 10 + b }
            __check((seq.joinToString(",")).toString(), "12,23,34,45")
        }
