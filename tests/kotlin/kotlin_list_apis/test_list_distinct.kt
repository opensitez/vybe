// vybe-test: kotlin/kotlin_list_apis/test_list_distinct
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(1, 1, 2, 2, 3)
            __check((list.distinct().joinToString(",")).toString(), "1,2,3")
        }
