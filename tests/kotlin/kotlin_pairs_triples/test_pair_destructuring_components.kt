// vybe-test: kotlin/kotlin_pairs_triples/test_pair_destructuring_components
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (left, right) = Pair("a", 9)
            __check((left).toString(), "a")
            __check((right).toString(), "9")
        }
