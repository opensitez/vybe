// vybe-test: kotlin/mutable_list_apis/test_mutable_list_plus_assign
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf(1)
            values += 2
            values += 3
            __check((values.joinToString(",")).toString(), "1,2,3")
        }
