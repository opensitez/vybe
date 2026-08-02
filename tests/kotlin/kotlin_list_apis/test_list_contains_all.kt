// vybe-test: kotlin/kotlin_list_apis/test_list_contains_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3, 4)
            __check((list.containsAll(listOf(1, 3))).toString(), "true")
            __check((list.containsAll(listOf(1, 7))).toString(), "false")
        }
