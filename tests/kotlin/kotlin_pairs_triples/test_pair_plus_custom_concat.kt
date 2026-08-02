// vybe-test: kotlin/kotlin_pairs_triples/test_pair_plus_custom_concat
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1 to 2)
            val b = listOf(3 to 4)
            val combined = a + b
            val out = combined.joinToString(",") { "${'$'}{it.first}=${'$'}{it.second}" }
            __check((out).toString(), "1=2,3=4")
        }
