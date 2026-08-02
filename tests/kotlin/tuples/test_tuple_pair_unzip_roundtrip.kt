// vybe-test: kotlin/tuples/test_tuple_pair_unzip_roundtrip
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = listOf(Pair(1, "a"), Pair(2, "b"), Pair(3, "c"))
            val (nums, chars) = source.unzip()
            __check((nums.joinToString(",")).toString(), "1,2,3")
            __check((chars.joinToString("")).toString(), "abc")
        }
