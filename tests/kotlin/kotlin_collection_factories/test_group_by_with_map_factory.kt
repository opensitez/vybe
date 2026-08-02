// vybe-test: kotlin/kotlin_collection_factories/test_group_by_with_map_factory
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4).groupBy { it % 2 == 0 }
            val even = values[true]?.sorted() ?: emptyList()
            val odd = values[false]?.sorted() ?: emptyList()
            __check((even.joinToString(",")).toString(), "2,4")
            __check((odd.joinToString(",")).toString(), "1,3")
        }
