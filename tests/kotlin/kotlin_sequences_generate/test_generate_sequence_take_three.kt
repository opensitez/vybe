// vybe-test: kotlin/kotlin_sequences_generate/test_generate_sequence_take_three
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { it + 1 }.take(3).toList()
            __check((seq.joinToString(",")).toString(), "1,2,3")
        }
