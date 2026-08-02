// vybe-test: kotlin/collections_sequences/test_generate_sequence_with_seed_and_limit
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { v -> if (v < 4) v + 1 else null }
            __check((seq.toList().joinToString(",")).toString(), "1,2,3,4")
        }
