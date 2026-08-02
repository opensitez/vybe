// vybe-test: kotlin/kotlin_collection_factories/test_linked_set_preserves_insertion_order_for_distinct_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = linkedSetOf("z", "a", "m")
            __check((values.joinToString(",")).toString(), "z,a,m")
        }
