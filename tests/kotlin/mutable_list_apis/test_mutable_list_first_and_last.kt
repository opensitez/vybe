// vybe-test: kotlin/mutable_list_apis/test_mutable_list_first_and_last
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(4, 5, 6)
            __check((values.first()).toString(), "4")
            __check((values.last()).toString(), "6")
        }
