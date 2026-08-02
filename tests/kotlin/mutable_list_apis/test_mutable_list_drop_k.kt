// vybe-test: kotlin/mutable_list_apis/test_mutable_list_drop_k
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2, 3, 4)
            val tail = values.drop(2)
            __check((tail.joinToString(",")).toString(), "3,4")
        }
