// vybe-test: kotlin/kotlin_list_apis/test_list_mutable_add_remove
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = mutableListOf(1, 2)
            list.add(3)
            list.add(1, 9)
            list.removeAt(0)
            __check((list.joinToString(",")).toString(), "9,2,3")
        }
