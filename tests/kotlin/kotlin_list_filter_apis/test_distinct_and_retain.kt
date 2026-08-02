// vybe-test: kotlin/kotlin_list_filter_apis/test_distinct_and_retain
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 2, 3, 3, 4)
            __check((nums.distinct().joinToString(",")).toString(), "1,2,3,4")
            val words = listOf("a", "ab", "ac", "b")
            __check((words.distinctBy { it[0] }.joinToString(",")).toString(), "a,b")
        }
