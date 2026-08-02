// vybe-test: kotlin/tuples/test_tuple_nested_tuple_unpacking_via_indexing
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = Pair(Triple(1, 2, 3), Triple(4, 5, 6))
            __check((nested.first.second + nested.second.first).toString(), "6")
        }
