// vybe-test: kotlin/type_aliases/test_typealias_for_pair_projection_shape
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias PairAlias = Pair<String, Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: PairAlias = Pair("x", 7)
            __check((value.first).toString(), "x")
            __check((value.second).toString(), "7")
        }
