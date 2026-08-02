// vybe-test: kotlin/kotlin_pairs_apis/test_pair_zip_basic
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val zipped = listOf(1, 2, 3).zip(listOf("a", "b", "c"))
            __check((zipped[0].first).toString(), "1")
            __check((zipped[1].second).toString(), "b")
            __check((zipped.joinToString("|") { it.toString() }).toString(), "(1, a)|(2, b)|(3, c)")
        }
