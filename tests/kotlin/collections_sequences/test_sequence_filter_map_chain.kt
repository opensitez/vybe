// vybe-test: kotlin/collections_sequences/test_sequence_filter_map_chain
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..6).asSequence()
                .map { it + 1 }
                .filter { it % 2 == 0 }
            __check((seq.toList().joinToString(",")).toString(), "2,4,6,8")
        }
