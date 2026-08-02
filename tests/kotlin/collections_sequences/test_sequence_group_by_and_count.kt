// vybe-test: kotlin/collections_sequences/test_sequence_group_by_and_count
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = listOf("cat", "dog", "cow", "deer").asSequence()
            val grouped = seq.groupBy { it.first() }
            val c = grouped["c"]?.size ?: 0
            val d = grouped["d"]?.size ?: 0
            __check((c).toString(), "2")
            __check((d).toString(), "2")
        }
