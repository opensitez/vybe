// vybe-test: kotlin/kotlin_list_apis/test_list_element_at
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(8, 9, 10)
            __check((list.elementAt(1)).toString(), "9")
            __check((list.elementAtOrNull(9) ?: "na").toString(), "na")
        }
