// vybe-test: kotlin/collections_iterables/test_list_first_or_null_and_last_or_null
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(9, 1, 5)
            __check((nums.firstOrNull { it > 10 } ?: "none").toString(), "none")
            __check((nums.lastOrNull { it < 3 } ?: "none").toString(), "1")
        }
