// vybe-test: kotlin/kotlin_list_apis/test_list_any_all_none
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3)
            __check((list.any { it > 2 }).toString(), "true")
            __check((list.all { it > 0 }).toString(), "true")
            __check((list.none { it < 0 }).toString(), "true")
        }
