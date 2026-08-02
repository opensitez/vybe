// vybe-test: kotlin/collections_set/test_set_min_max_aggregation
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(5, 1, 9, 3)
            __check((values.minOrNull()).toString(), "1")
            __check((values.maxOrNull()).toString(), "9")
        }
