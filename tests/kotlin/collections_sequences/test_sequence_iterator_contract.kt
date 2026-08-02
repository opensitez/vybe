// vybe-test: kotlin/collections_sequences/test_sequence_iterator_contract
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2)
            val it = seq.iterator()
            __check((it.hasNext()).toString(), "true")
            __check((it.next()).toString(), "1")
            __check((it.hasNext()).toString(), "true")
            __check((it.next()).toString(), "2")
            __check((it.hasNext()).toString(), "false")
        }
