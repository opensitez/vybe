// vybe-test: kotlin/tuples/test_tuple_nested_pair_unwrap_access
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = Pair(Pair(1, 2), Pair(3, 4))
            __check((nested.first.first).toString(), "1")
            __check((nested.second.second).toString(), "4")
        }
