// vybe-test: kotlin/mutable_list_apis/test_mutable_list_add_and_get
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1, 2)
            values.add(3)
            __check((values[2]).toString(), "3")
            __check((values.size).toString(), "3")
        }
