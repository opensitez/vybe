// vybe-test: kotlin/mutable_list_apis/test_mutable_list_add_all_from_list
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1)
            values.addAll(listOf(2, 3))
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
