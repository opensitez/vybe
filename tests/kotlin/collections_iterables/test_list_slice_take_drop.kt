// vybe-test: kotlin/collections_iterables/test_list_slice_take_drop
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = listOf(1, 2, 3, 4, 5)
            __check((nums.take(3).joinToString(",")).toString(), "1,2,3")
            __check((nums.drop(2).joinToString(",")).toString(), "3,4,5")
            __check((nums.slice(1..3).joinToString(",")).toString(), "2,3,4")
        }
