// vybe-test: kotlin/kotlin_list_apis/test_list_first_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf("a", "b", "c")
            __check((list.first()).toString(), "a")
            __check((list.last()).toString(), "c")
        }
