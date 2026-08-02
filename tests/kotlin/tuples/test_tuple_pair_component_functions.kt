// vybe-test: kotlin/tuples/test_tuple_pair_component_functions
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair = Pair(8, "v")
            __check((pair.component1()).toString(), "8")
            __check((pair.component2()).toString(), "v")
        }
