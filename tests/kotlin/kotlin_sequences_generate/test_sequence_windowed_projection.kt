// vybe-test: kotlin/kotlin_sequences_generate/test_sequence_windowed_projection
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = generateSequence(0) { if (it == 3) null else it + 1 }
                .windowed(2)
                .toList()
            __check((values.size).toString(), "3")
            __check((values.joinToString("|") { it.joinToString("") }).toString(), "01|12|23")
        }
