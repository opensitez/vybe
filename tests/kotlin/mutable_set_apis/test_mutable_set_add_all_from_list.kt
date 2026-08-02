// vybe-test: kotlin/mutable_set_apis/test_mutable_set_add_all_from_list
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1)
            values.addAll(listOf(1, 2, 3))
            __check((values.joinToString(",")).toString(), "1,2,3")
            __check((values.size).toString(), "3")
        }
