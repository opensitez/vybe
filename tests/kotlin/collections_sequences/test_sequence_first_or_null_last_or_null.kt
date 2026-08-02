// vybe-test: kotlin/collections_sequences/test_sequence_first_or_null_last_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(9, 10, 11).asSequence()
            __check((seq.firstOrNull { it > 100 } ?: "none").toString(), "none")
            __check((seq.lastOrNull { it < 10 } ?: "none").toString(), "9")
        }
