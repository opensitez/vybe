// vybe-test: kotlin/mutable_set_apis/test_mutable_set_keep_distinct_order_after_mutation
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2)
            values.add(2)
            values.add(3)
            values.remove(1)
            __check((values.joinToString(",")).toString(), "2,3")
        }
