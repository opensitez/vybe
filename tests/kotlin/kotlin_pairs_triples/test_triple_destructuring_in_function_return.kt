// vybe-test: kotlin/kotlin_pairs_triples/test_triple_destructuring_in_function_return
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_triples.rs

fun make(): Triple<String, Int, Boolean> {
            return Triple("x", 4, true)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (k, n, b) = make()
            __check((k).toString(), "x")
            __check((n).toString(), "4")
            __check((b).toString(), "true")
        }
