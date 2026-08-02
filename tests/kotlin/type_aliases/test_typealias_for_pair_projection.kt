// vybe-test: kotlin/type_aliases/test_typealias_for_pair_projection
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias PairLike = Pair<Int, Boolean>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: PairLike = Pair(4, true)
            __check((value.first).toString(), "4")
            __check((value.second).toString(), "true")
        }
