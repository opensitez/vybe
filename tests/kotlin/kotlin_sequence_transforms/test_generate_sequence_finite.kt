// vybe-test: kotlin/kotlin_sequence_transforms/test_generate_sequence_finite
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { if (it < 4) it + 1 else null }
            __check((seq.toList().joinToString(",")).toString(), "1,2,3,4")
        }
