// vybe-test: kotlin/collections_sequences/test_sequence_filter_not_null_skips_nulls
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, null, 2, null, 3).filterNotNull()
            __check((seq.toList().joinToString(",")).toString(), "1,2,3")
        }
