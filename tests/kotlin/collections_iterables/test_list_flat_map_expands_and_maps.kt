// vybe-test: kotlin/collections_iterables/test_list_flat_map_expands_and_maps
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val groups = listOf(
                listOf(1, 2),
                listOf(3, 4)
            )
            val expanded = groups.flatMap { it }
            __check((expanded.joinToString(",")).toString(), "1,2,3,4")
            val mapped = groups.flatMap { inner -> inner.map { it * 10 } }
            __check((mapped.joinToString(",")).toString(), "10,20,30,40")
        }
