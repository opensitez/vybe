// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_map_and_reduce
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = sequenceOf(1, 2, 3)
                .map { it + 1 }
                .reduce { acc, value -> acc + value }
            __check((out).toString(), "9")
        }
