// vybe-test: kotlin/kotlin_pairs_apis/test_pair_inequality
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 1 to 2
            val b = 2 to 1
            __check((a == b).toString(), "false")
        }
