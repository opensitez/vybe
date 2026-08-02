// vybe-test: kotlin/kotlin_collection_chunking/test_flat_map_lists
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_chunking.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(listOf(1, 2), listOf(3, 4))
            val mapped = source.flatMap { it.map { v -> v * 10 } }
            __check((mapped.joinToString(",")).toString(), "10,20,30,40")
        }
