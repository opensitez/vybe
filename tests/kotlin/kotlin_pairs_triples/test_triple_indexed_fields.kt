// vybe-test: kotlin/kotlin_pairs_triples/test_triple_indexed_fields
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Triple(1, 2, 3)
            __check((t.third).toString(), "3")
            __check((t.toList().joinToString(",")).toString(), "1,2,3")
        }
