// vybe-test: kotlin/collections_sequences/test_sequence_single_and_single_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seqOne = listOf(7).asSequence()
            val seqEmpty = emptySequence<Int>()
            __check((seqOne.single()).toString(), "7")
            __check((seqEmpty.singleOrNull() ?: -1).toString(), "-1")
        }
