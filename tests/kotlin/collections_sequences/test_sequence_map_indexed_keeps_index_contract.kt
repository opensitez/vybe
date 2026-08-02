// vybe-test: kotlin/collections_sequences/test_sequence_map_indexed_keeps_index_contract
// origin: languages/kotlin/tests/kotlin/test_collections_sequences.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val labeled = listOf("x", "y", "z").asSequence()
                .mapIndexed { index, value -> "$index:$value" }
                .toList()
            __check((labeled.joinToString(",")).toString(), "0:x,1:y,2:z")
        }
