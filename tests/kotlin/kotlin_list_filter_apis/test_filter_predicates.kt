// vybe-test: kotlin/kotlin_list_filter_apis/test_filter_predicates
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.filter { it % 2 == 0 }.joinToString(",")).toString(), "2,4")
            __check((nums.filterNot { it > 3 }.joinToString(",")).toString(), "1,2,3")
            __check((nums.filterIndexed { index, _ -> index % 2 == 0 }.joinToString(",")).toString(), "1,3,5")
        }
