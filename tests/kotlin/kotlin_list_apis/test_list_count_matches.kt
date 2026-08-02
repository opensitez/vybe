// vybe-test: kotlin/kotlin_list_apis/test_list_count_matches
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3, 4, 5)
            __check((list.count { it % 2 == 0 }).toString(), "2")
            __check((list.count { it > 4 }).toString(), "1")
        }
