// vybe-test: kotlin/kotlin_list_apis/test_list_sorted_descending
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(3, 1, 4, 2)
            __check((list.sortedDescending().joinToString(",")).toString(), "4,3,2,1")
        }
