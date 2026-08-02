// vybe-test: kotlin/collections_iterables/test_reversed_list_is_independent_snapshot
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = mutableListOf(1, 2, 3)
            val snapshot = nums.reversed()
            nums[0] = 9
            nums.add(4)
            __check((snapshot.joinToString(",")).toString(), "3,2,1")
            __check((nums.joinToString(",")).toString(), "9,2,3,4")
        }
