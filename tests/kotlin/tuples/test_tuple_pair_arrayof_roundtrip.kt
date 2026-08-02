// vybe-test: kotlin/tuples/test_tuple_pair_arrayof_roundtrip
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = arrayOf(Pair(9, 8), Pair(7, 6))
            val list = source.toList()
            __check((list[0].first + list[1].second).toString(), "15")
            __check((list[1].toString()).toString(), "(7, 6)")
        }
