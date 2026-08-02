// vybe-test: kotlin/kotlin_list_filter_apis/test_drop_while_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3)
            __check((nums.dropWhile { it < 0 }.joinToString(",")).toString(), "1,2,3")
            __check((nums.takeWhile { false }.joinToString(",")).toString(), "")
            __check((nums.takeWhile { true }.size).toString(), "3")
        }
