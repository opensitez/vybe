// vybe-test: kotlin/mutable_list_apis/test_mutable_list_reverse_in_place
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3)
            values.reverse()
            __check((values.joinToString(",")).toString(), "3,2,1")
        }
