// vybe-test: kotlin/collections_iterables/test_element_at_and_element_at_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(10, 20, 30)
            __check((nums.elementAt(1)).toString(), "20")
            __check((nums.elementAtOrNull(5) ?: -1).toString(), "-1")
        }
