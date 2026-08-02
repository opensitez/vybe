// vybe-test: kotlin/kotlin_pairs_apis/test_pair_unzip_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (a, b) = emptyList<Pair<Int, Int>>().unzip()
            __check((a.isEmpty()).toString(), "true")
            __check((b.isEmpty()).toString(), "true")
        }
