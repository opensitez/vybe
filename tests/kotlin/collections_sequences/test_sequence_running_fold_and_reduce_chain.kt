// vybe-test: kotlin/collections_sequences/test_sequence_running_fold_and_reduce_chain
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4).runningFold(0) { acc, value -> acc + value }.toList()
            __check((seq.joinToString(",")).toString(), "0,1,3,6,10")
            val running = sequenceOf(1, 2, 3).runningReduce { acc, value -> acc * value }.toList()
            __check((running.joinToString(",")).toString(), "1,2,6")
        }
