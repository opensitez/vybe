// vybe-test: kotlin/collections_sequences/test_sequence_fold_reduce
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf(1, 2, 3, 4).asSequence()
            __check((seq.fold(0) { acc, n -> acc + n }).toString(), "10")
            __check((listOf(4, 3, 2).asSequence().reduce { acc, n -> acc - n }).toString(), "-1")
        }
