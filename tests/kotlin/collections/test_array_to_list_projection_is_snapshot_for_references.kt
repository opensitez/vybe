// vybe-test: kotlin/collections/test_array_to_list_projection_is_snapshot_for_references
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = IntArray(3) { it + 1 }
            val snapshot = nums.toList()
            nums[0] = 9
            __check((snapshot.joinToString(",")).toString(), "1,2,3")
            __check((nums.joinToString(",")).toString(), "9,2,3")
            __check((snapshot[1]).toString(), "2")
        }
