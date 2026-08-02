// vybe-test: kotlin/tuples/test_tuple_pair_of_lists_projection_to_map_key_values
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val rows = listOf(Pair("a", 1), Pair("b", 2), Pair("c", 3))
            val labels = rows.toMap()
            __check((labels["b"]).toString(), "2")
            __check((labels["x"] ?: -1).toString(), "-1")
            __check((labels.size).toString(), "3")
        }
