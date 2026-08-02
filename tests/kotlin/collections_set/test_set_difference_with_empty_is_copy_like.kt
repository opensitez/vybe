// vybe-test: kotlin/collections_set/test_set_difference_with_empty_is_copy_like
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            val remaining = values - emptySet<Int>()
            __check((remaining.size).toString(), "3")
            __check((remaining == values).toString(), "true")
        }
