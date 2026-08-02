// vybe-test: kotlin/collections_iterables/test_list_flatten_nested_structure
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = listOf(
                listOf("x", "y"),
                listOf(),
                listOf("z")
            )
            __check((nested.flatten().joinToString(",")).toString(), "x,y,z")
        }
