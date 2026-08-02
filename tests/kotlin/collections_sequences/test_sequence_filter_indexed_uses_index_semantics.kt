// vybe-test: kotlin/collections_sequences/test_sequence_filter_indexed_uses_index_semantics
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = (10..16).asSequence()
                .filterIndexed { index, value -> index % 2 == 0 && value > 11 }
                .toList()
                .joinToString(",")
            __check((out).toString(), "14")
        }
