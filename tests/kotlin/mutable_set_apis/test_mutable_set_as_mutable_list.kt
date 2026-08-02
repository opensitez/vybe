// vybe-test: kotlin/mutable_set_apis/test_mutable_set_as_mutable_list
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            val list = values.toMutableList()
            list.sort()
            __check((list.joinToString(",")).toString(), "1,2,3")
        }
