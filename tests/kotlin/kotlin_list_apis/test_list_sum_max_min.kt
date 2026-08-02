// vybe-test: kotlin/kotlin_list_apis/test_list_sum_max_min
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(3, 9, 1, 5)
            __check((list.sum()).toString(), "18")
            __check((list.maxOrNull()).toString(), "9")
            __check((list.minOrNull()).toString(), "1")
        }
