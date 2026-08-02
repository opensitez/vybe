// vybe-test: kotlin/kotlin_pairs_apis/test_pair_and_triple_nested
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested = Pair("outer", Triple(1, "mid", 3))
            val (left, right) = nested
            __check((left).toString(), "outer")
            __check((right.second).toString(), "mid")
        }
