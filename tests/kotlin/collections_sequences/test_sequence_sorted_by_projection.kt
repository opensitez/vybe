// vybe-test: kotlin/collections_sequences/test_sequence_sorted_by_projection
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf("aa", "b", "ccc").asSequence()
            __check((source.sortedBy { it.length }.joinToString(",")).toString(), "b,aa,ccc")
        }
