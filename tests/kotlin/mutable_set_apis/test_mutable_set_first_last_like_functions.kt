// vybe-test: kotlin/mutable_set_apis/test_mutable_set_first_last_like_functions
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(2, 4, 6)
            __check((values.first()).toString(), "2")
            __check((values.last()).toString(), "6")
        }
