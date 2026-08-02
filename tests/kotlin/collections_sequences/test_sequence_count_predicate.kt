// vybe-test: kotlin/collections_sequences/test_sequence_count_predicate
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 2, 3, 4, 5, 6).asSequence()
            __check((seq.count { it % 3 == 0 }).toString(), "2")
            __check((seq.count()).toString(), "6")
        }
