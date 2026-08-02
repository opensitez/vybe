// vybe-test: kotlin/collections_sequences/test_sequence_take_and_drop
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..10).asSequence()
            __check((seq.take(4).toList().joinToString(",")).toString(), "1,2,3,4")
            __check(((1..10).asSequence().drop(7).toList().joinToString(",")).toString(), "8,9,10")
        }
