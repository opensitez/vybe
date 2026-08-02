// vybe-test: kotlin/kotlin_list_filter_apis/test_take_drop_operations
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5, 6)
            __check((nums.take(3).joinToString(",")).toString(), "1,2,3")
            __check((nums.drop(3).joinToString(",")).toString(), "4,5,6")
            __check((nums.takeLast(2).joinToString(",")).toString(), "5,6")
            __check((nums.dropLast(2).joinToString(",")).toString(), "1,2,3,4")
        }
