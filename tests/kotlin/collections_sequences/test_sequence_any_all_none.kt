// vybe-test: kotlin/collections_sequences/test_sequence_any_all_none
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(2, 4, 6).asSequence()
            __check((seq.any { it > 5 }).toString(), "true")
            __check((seq.all { it % 2 == 0 }).toString(), "true")
            __check((seq.none { it == 9 }).toString(), "true")
        }
