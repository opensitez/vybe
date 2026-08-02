// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_replayability
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { if (it < 3) it + 1 else null }
            __check((seq.toList().size).toString(), "3")
            // recreate because sequences are consumed
            val seq2 = generateSequence(1) { if (it < 3) it + 1 else null }
            __check((seq2.toList().sum().toString()).toString(), "6")
        }
