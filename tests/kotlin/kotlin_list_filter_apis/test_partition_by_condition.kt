// vybe-test: kotlin/kotlin_list_filter_apis/test_partition_by_condition
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(5, 10, 11, 20)
            val (evens, odds) = nums.partition { it % 2 == 0 }
            __check((evens.joinToString(",")).toString(), "10,20")
            __check((odds.joinToString(",")).toString(), "5,11")
            __check((evens.count()).toString(), "2")
        }
