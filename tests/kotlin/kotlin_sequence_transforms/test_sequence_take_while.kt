// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_take_while
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf(1, 2, 3, 0, 4)
            __check((seq.takeWhile { it > 0 }.toList().joinToString(",")).toString(), "1,2,3")
            __check((seq.dropWhile { it < 3 }.toList().joinToString(",")).toString(), "3,0,4")
        }
