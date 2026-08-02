// vybe-test: kotlin/kotlin_list_apis/test_list_index_of
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = listOf(10, 20, 30, 20)
            __check((list.indexOf(20)).toString(), "1")
            __check((list.lastIndexOf(20)).toString(), "3")
            __check((list.indexOf(99)).toString(), "-1")
        }
