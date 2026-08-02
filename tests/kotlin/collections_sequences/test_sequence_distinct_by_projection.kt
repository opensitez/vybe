// vybe-test: kotlin/collections_sequences/test_sequence_distinct_by_projection
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf("aa", "ab", "b", "cc").asSequence().distinctBy { it.length }
            __check((seq.toList().joinToString("|")).toString(), "aa,b")
        }
