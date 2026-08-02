// vybe-test: kotlin/kotlin_list_filter_apis/test_filter_with_index_and_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(10, 11, 12, 13)
            val filtered = nums.filterIndexed { idx, value -> idx + value == 11 || idx + value == 15 }
            __check((filtered.joinToString(",")).toString(), "10,12")
        }
