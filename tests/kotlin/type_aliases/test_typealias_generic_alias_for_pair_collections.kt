// vybe-test: kotlin/type_aliases/test_typealias_generic_alias_for_pair_collections
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias PairList<T> = List<Pair<T, T>>

        fun total(values: PairList<Int>): Int {
            return values.fold(0) { acc, item -> acc + item.first + item.second }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: PairList<Int> = listOf(Pair(1, 2), Pair(3, 4))
            __check((total(values)).toString(), "10")
        }
