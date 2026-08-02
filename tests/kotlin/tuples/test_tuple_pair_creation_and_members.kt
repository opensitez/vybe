// vybe-test: kotlin/tuples/test_tuple_pair_creation_and_members
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair("x", 4)
            __check((pair.first).toString(), "x")
            __check((pair.second).toString(), "4")
        }
