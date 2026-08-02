// vybe-test: kotlin/collections/test_array_swap_in_place
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(1, 2, 3, 4)
            val tmp = nums[0]
            nums[0] = nums[3]
            nums[3] = tmp
            __check((nums[0] + nums[3]).toString(), "5")
            __check((nums[1] + nums[2]).toString(), "5")
        }
