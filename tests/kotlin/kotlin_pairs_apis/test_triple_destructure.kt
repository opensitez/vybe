// vybe-test: kotlin/kotlin_pairs_apis/test_triple_destructure
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (a, b, c) = Triple(2, 4, 6)
            __check((a).toString(), "2")
            __check((b).toString(), "4")
            __check((c).toString(), "6")
        }
