// vybe-test: kotlin/kotlin_collection_factories/test_sorted_set_orders_lexicographically
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = sortedSetOf(5, 1, 4, 2, 3)
            __check((values.joinToString(",")).toString(), "1,2,3,4,5")
            __check((values.first()).toString(), "1")
            __check((values.last()).toString(), "5")
        }
