// vybe-test: kotlin/collections_iterables/test_list_drop_last_take_last
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.dropLast(2).joinToString(",")).toString(), "1,2,3")
            __check((nums.takeLast(3).joinToString(",")).toString(), "3,4,5")
        }
