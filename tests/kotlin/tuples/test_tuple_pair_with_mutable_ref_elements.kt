// vybe-test: kotlin/tuples/test_tuple_pair_with_mutable_ref_elements
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mutableListOf(1)
            val b = mutableListOf(2)
            val pair = Pair(a, b)
            pair.first.add(3)
            pair.second.add(4)
            __check((a.size).toString(), "2")
            __check((b.size).toString(), "2")
            __check((pair.first[1] + pair.second[1]).toString(), "7")
        }
