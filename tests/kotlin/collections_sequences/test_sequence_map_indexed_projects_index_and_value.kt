// vybe-test: kotlin/collections_sequences/test_sequence_map_indexed_projects_index_and_value
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = sequenceOf("a", "bb", "ccc")
            val annotated = seq.mapIndexed { index, value ->
                value.length + index
            }.toList().joinToString(",")
            __check((annotated).toString(), "1,3,5")
        }
