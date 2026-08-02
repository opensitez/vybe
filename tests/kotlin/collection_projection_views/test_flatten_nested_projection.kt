// vybe-test: kotlin/collection_projection_views/test_flatten_nested_projection
// origin: languages/kotlin/tests/kotlin/test_collection_projection_views.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = listOf(listOf(1, 2), listOf(3), listOf(4, 5))
            __check((nested.flatten().joinToString(",")).toString(), "1,2,3,4,5")
            __check((nested.flatMap { it.map { v -> v * 2 } }.joinToString(",")).toString(), "2,4,6,8,10")
        }
