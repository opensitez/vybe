// vybe-test: kotlin/kotlin_pairs_triples/test_triple_component_functions
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val point = Triple(1, 2, 3)
            __check((point.component1()).toString(), "1")
            __check((point.component2()).toString(), "2")
            __check((point.component3()).toString(), "3")
        }
