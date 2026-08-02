// vybe-test: kotlin/kotlin_list_apis/test_list_take_last_drop_last
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 2, 3, 4)
            __check((list.takeLast(2).joinToString(",")).toString(), "3,4")
            __check((list.dropLast(2).joinToString(",")).toString(), "1,2")
        }
