// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_chunked_with_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sum = (1..6).asSequence().chunked(2) { it.sum() }.toList()
            __check((sum.joinToString(",")).toString(), "3,7,11")
        }
