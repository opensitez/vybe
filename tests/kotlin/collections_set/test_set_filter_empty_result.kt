// vybe-test: kotlin/collections_set/test_set_filter_empty_result
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            val none = values.filter { it > 10 }.toSet()
            __check((none.isEmpty()).toString(), "true")
            __check((none.size).toString(), "0")
        }
