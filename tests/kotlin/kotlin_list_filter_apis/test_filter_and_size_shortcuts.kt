// vybe-test: kotlin/kotlin_list_filter_apis/test_filter_and_size_shortcuts
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.all { it > 0 }).toString(), "true")
            __check((nums.any { it > 4 }).toString(), "true")
            __check((nums.none { it > 10 }).toString(), "true")
        }
