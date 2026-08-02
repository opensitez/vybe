// vybe-test: kotlin/collections_sequences/test_sequence_take_while_drop_while
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 3, 5, 2, 4, 6).asSequence()
            __check((seq.takeWhile { it < 4 }.toList().joinToString(",")).toString(), "1,3")
            __check((seq.dropWhile { it < 4 }.toList().joinToString(",")).toString(), "5,2,4,6")
        }
