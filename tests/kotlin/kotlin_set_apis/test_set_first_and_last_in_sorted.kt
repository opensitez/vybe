// vybe-test: kotlin/kotlin_set_apis/test_set_first_and_last_in_sorted
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = sortedSetOf(9, 1, 4)
            __check((set.first()).toString(), "1")
            __check((set.last()).toString(), "9")
        }
