// vybe-test: kotlin/collections_iterables/test_list_contains_all_and_sublist
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.containsAll(listOf(2, 4))).toString(), "true")
            __check((nums.subList(1, 4).joinToString(",")).toString(), "2,3,4")
        }
