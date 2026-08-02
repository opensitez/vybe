// vybe-test: kotlin/kotlin_sequences_generate/test_as_sequence_projection_chain
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = (1..6).asSequence().filter { it % 2 == 0 }.map { it * it }.take(2).toList()
            __check((out.joinToString(",")).toString(), "4,16")
        }
