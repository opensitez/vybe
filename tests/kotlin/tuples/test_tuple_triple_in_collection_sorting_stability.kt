// vybe-test: kotlin/tuples/test_tuple_triple_in_collection_sorting_stability
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val points = listOf(
                Triple("a", 3, 2),
                Triple("b", 1, 9)
            )
            __check((points.sortedBy { it.second }.joinToString("|") { it.first }).toString(), "b|a")
        }
