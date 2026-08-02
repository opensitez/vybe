// vybe-test: kotlin/kotlin_pairs_apis/test_pair_infix_constructor
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = "left" to "right"
            __check((value.first).toString(), "left")
            __check((value.second).toString(), "right")
        }
