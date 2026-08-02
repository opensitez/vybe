// vybe-test: kotlin/ordered_collections/test_collection_concatenation_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf(1, 2)
            val b = listOf(3, 4)
            __check(((a + b).joinToString(",")).toString(), "1,2,3,4")
            __check((a.plus(b).joinToString(",")).toString(), "1,2,3,4")
        }
