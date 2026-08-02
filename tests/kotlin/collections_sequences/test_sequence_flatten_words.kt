// vybe-test: kotlin/collections_sequences/test_sequence_flatten_words
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(listOf("a"), emptyList(), listOf("b", "c")).asSequence().flatten()
            __check((seq.joinToString(",")).toString(), "a,b,c")
        }
