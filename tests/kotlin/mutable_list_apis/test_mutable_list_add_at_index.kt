// vybe-test: kotlin/mutable_list_apis/test_mutable_list_add_at_index
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 3, 4)
            values.add(1, 2)
            __check((values.joinToString(",")).toString(), "1,2,3,4")
        }
