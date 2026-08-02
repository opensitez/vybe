// vybe-test: kotlin/collections_sequences/test_sequence_windowed_with_partial_windows
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = (1..5).asSequence()
            __check((source.windowed(2, 2, partialWindows = true).joinToString("|") { it.joinToString("-") }).toString(), "1-2|3-4|5")
            __check((source.windowed(3).toList().isEmpty()).toString(), "false")
        }
