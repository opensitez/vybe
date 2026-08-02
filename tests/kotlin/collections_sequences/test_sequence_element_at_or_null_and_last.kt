// vybe-test: kotlin/collections_sequences/test_sequence_element_at_or_null_and_last
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf("x", "y", "z").asSequence()
            __check((seq.elementAt(1)).toString(), "y")
            __check((seq.elementAtOrNull(5) ?: "none").toString(), "none")
            __check((seq.last()).toString(), "z")
        }
