// vybe-test: kotlin/kotlin_pairs_apis/test_pair_of_empty_string
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = "" to 0
            __check((p.first.isEmpty()).toString(), "true")
            __check((p.second).toString(), "0")
        }
