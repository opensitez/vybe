// vybe-test: kotlin/kotlin_list_apis/test_list_first_or_null
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1)
            __check((list.firstOrNull()).toString(), "1")
            __check((emptyList<Int>().firstOrNull() ?: "none").toString(), "none")
        }
