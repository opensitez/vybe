// vybe-test: kotlin/kotlin_list_apis/test_list_find_first
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(5, 6, 7)
            __check((list.find { it > 5 }).toString(), "6")
            __check((list.findLast { it < 6 } ?: "none").toString(), "5")
        }
