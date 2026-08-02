// vybe-test: kotlin/kotlin_pairs_triples/test_pair_to_infix_constructor
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = 4 to "four"
            __check((p.first + 1).toString(), "5")
            __check((p.second.length).toString(), "4")
        }
