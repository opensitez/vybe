// vybe-test: kotlin/type_aliases/test_typealias_for_tuple_like_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias PairText = Pair<Int, String>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: PairText = Pair(4, "x")
            __check((value.first).toString(), "4")
            __check((value.second).toString(), "x")
        }
