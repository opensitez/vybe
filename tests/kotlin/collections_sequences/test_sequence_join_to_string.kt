// vybe-test: kotlin/collections_sequences/test_sequence_join_to_string
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..4).asSequence()
            __check((seq.joinToString(prefix = "[", postfix = "]")).toString(), "[1, 2, 3, 4]")
        }
