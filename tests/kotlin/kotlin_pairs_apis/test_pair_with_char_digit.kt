// vybe-test: kotlin/kotlin_pairs_apis/test_pair_with_char_digit
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = 'A' to 5
            __check((value.first.code).toString(), "65")
            __check((value.second).toString(), "5")
            __check((value.toString()).toString(), "(A, 5)")
        }
