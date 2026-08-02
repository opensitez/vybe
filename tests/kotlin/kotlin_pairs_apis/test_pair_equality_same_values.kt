// vybe-test: kotlin/kotlin_pairs_apis/test_pair_equality_same_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 4 to 5
            val b = 4 to 5
            __check((a == b).toString(), "true")
            __check((a != b).toString(), "false")
        }
