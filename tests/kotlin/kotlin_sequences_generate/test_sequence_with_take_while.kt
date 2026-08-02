// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_with_take_while
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = generateSequence(1) { it + 2 }
                .takeWhile { it < 8 }
                .toList()
            __check((out.joinToString(",")).toString(), "1,3,5,7")
        }
