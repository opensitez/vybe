// vybe-test: kotlin/collections_sequences/test_empty_sequence_has_expected_behaviors
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = emptySequence<Int>()
            __check((seq.count()).toString(), "0")
            __check((seq.none { true }).toString(), "true")
            __check((seq.toList().size).toString(), "0")
        }
