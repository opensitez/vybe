// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_count_take
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..10).asSequence()
            __check((seq.count()).toString(), "10")
            __check((seq.take(3).toList().size).toString(), "3")
        }
