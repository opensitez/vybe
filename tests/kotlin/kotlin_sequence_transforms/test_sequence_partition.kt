// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_partition
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 4, 5)
            val (evens, odds) = seq.partition { it % 2 == 0 }
            __check((evens.toList().joinToString(",")).toString(), "2,4")
            __check((odds.toList().joinToString(",")).toString(), "1,3,5")
        }
