// vybe-test: kotlin/collections_sequences/test_sequence_find_first_last
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(5, 12, 17, 20).asSequence()
            __check((seq.find { it % 2 == 0 } ?: -1).toString(), "12")
            __check((seq.findLast { it % 2 == 1 } ?: -1).toString(), "17")
        }
