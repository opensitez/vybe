// vybe-test: kotlin/kotlin_list_filter_apis/test_slice_window_chunk
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.slice(1..3).joinToString(",")).toString(), "2,3,4")
            __check((nums.subList(0, 2).joinToString(",")).toString(), "1,2")
            __check((nums.sliceArray(2 until 4).joinToString(",")).toString(), "3,4")
        }
