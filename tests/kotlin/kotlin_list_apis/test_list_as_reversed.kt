// vybe-test: kotlin/kotlin_list_apis/test_list_as_reversed
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(7, 8, 9)
            __check((list.asReversed().joinToString(",")).toString(), "9,8,7")
            __check((list.joinToString(",")).toString(), "7,8,9")
        }
