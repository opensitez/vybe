// vybe-test: kotlin/collections_iterables/test_mutable_list_operations_modify_size_and_order
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = mutableListOf(1, 2, 3)
            nums.add(4)
            nums.removeAt(1)
            nums[0] = 8
            nums.remove(3)
            __check((nums.joinToString(",")).toString(), "8,4")
            __check((nums.size).toString(), "2")
        }
