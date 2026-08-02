// vybe-test: kotlin/tuples/test_triple_used_as_pair_like_in_map_with_projection
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val points = listOf(
                Triple("a", 1, 10),
                Triple("b", 2, 20)
            ).associateBy { it.first }
            val first = "a"
            val values = points[first]!!
            __check((first).toString(), "a")
            __check((values.second).toString(), "1")
            __check((values.third).toString(), "10")
        }
