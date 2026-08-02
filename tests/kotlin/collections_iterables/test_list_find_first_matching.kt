// vybe-test: kotlin/collections_iterables/test_list_find_first_matching
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(5, 12, 13, 20)
            __check((nums.find { it % 2 == 0 }).toString(), "12")
            __check((nums.findLast { it % 2 == 1 }).toString(), "13")
        }
