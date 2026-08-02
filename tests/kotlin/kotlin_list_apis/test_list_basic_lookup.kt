// vybe-test: kotlin/kotlin_list_apis/test_list_basic_lookup
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(4, 5, 6)
            __check((list[0]).toString(), "4")
            __check((list.size).toString(), "3")
            __check((list[2]).toString(), "6")
        }
