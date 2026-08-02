// vybe-test: kotlin/collections_iterables/test_list_any_and_all_and_none
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4)
            __check((nums.any { it > 3 }).toString(), "true")
            __check((nums.all { it < 10 }).toString(), "true")
            __check((nums.none { it == 9 }).toString(), "true")
        }
