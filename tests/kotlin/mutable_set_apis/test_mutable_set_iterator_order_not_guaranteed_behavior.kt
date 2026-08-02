// vybe-test: kotlin/mutable_set_apis/test_mutable_set_iterator_order_not_guaranteed_behavior
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(3, 1, 2)
            val items = values.toMutableList()
            items.sort()
            __check((items.joinToString(",")).toString(), "1,2,3")
            __check((items.size).toString(), "3")
        }
