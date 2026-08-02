// vybe-test: kotlin/collections_sequences/test_sequence_with_chunked_windows
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = (1..5).asSequence()
            __check((seq.chunked(2).joinToString("|") { it.joinToString("-") }).toString(), "1-2|3-4|5")
            __check(((1..5).asSequence().windowed(3).joinToString("|") { it.joinToString("-") }).toString(), "1-2-3|2-3-4|3-4-5")
        }
