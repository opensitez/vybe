// vybe-test: kotlin/collections_iterables/test_list_take_while_and_drop_while
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 3, 5, 2, 4, 6, 8)
            __check((nums.takeWhile { it < 4 }.joinToString(",")).toString(), "1,3")
            __check((nums.dropWhile { it < 4 }.joinToString(",")).toString(), "5,2,4,6,8")
        }
