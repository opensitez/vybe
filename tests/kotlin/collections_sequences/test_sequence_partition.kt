// vybe-test: kotlin/collections_sequences/test_sequence_partition
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..6).asSequence()
            val (lt4, ge4) = seq.partition { it < 4 }
            __check((lt4.joinToString(",")).toString(), "1,2,3")
            __check((ge4.joinToString(",")).toString(), "4,5,6")
        }
