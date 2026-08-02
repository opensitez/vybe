// vybe-test: kotlin/kotlin_list_apis/test_list_retain_if_even
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = mutableListOf(1, 2, 3, 4, 5)
            list.retainAll { it % 2 == 0 }
            __check((list.joinToString(",")).toString(), "2,4")
        }
