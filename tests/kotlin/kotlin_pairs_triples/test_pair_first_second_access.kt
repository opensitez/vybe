// vybe-test: kotlin/kotlin_pairs_triples/test_pair_first_second_access
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Pair(10, "ok")
            __check((p.first).toString(), "10")
            __check((p.second).toString(), "ok")
        }
