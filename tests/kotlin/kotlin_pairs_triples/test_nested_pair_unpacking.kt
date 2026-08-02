// vybe-test: kotlin/kotlin_pairs_triples/test_nested_pair_unpacking
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val outer = Pair(Pair(1, 2), Pair(3, 4))
            val (left, right) = outer
            val (a, b) = left
            val (c, d) = right
            __check((a + b + c + d).toString(), "10")
        }
