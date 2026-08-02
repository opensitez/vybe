// vybe-test: kotlin/kotlin_collection_factories/test_flatten_with_map_of_lists
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = mapOf("a" to listOf(1, 2), "b" to listOf(3), "c" to emptyList())
            val flat = value.values.flatten()
            __check((flat.joinToString(",")).toString(), "1,2,3")
            __check((flat.size).toString(), "3")
        }
