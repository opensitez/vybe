// vybe-test: kotlin/kotlin_list_filter_apis/test_take_while_drop_while
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 2, 1)
            __check((nums.takeWhile { it < 3 }.joinToString(",")).toString(), "1,2")
            __check((nums.dropWhile { it < 3 }.joinToString(",")).toString(), "3,2,1")
            __check((nums.takeWhile { it > 0 }.joinToString(",")).toString(), "1,2,3,2,1")
        }
