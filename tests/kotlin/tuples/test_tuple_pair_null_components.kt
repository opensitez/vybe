// vybe-test: kotlin/tuples/test_tuple_pair_null_components
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair: Pair<String?, Int?> = Pair(null, null)
            __check((pair.first == null).toString(), "true")
            __check((pair.second == null).toString(), "true")
        }
