// vybe-test: kotlin/kotlin_list_apis/test_list_join_to_string_custom
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3)
            __check((list.joinToString(prefix = "[", postfix = "]", separator = ":")).toString(), "[1:2:3]")
        }
