// vybe-test: kotlin/ordered_collections/test_sequence_of_list_indices
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            val indices = values.indices.toList()
            __check((indices.joinToString(",")).toString(), "0,1,2")
            __check((values[indices.first()]).toString(), "a")
            __check((values[indices.last()]).toString(), "c")
        }
