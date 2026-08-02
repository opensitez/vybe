// vybe-test: kotlin/kotlin_pairs_apis/test_pair_destructure
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (first, second) = "a" to "b"
            __check((first).toString(), "a")
            __check((second).toString(), "b")
        }
