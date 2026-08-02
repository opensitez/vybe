// vybe-test: kotlin/collections/test_array_mutation_in_place
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(10, 20, 30)
            nums[1] = 99
            __check((nums[1]).toString(), "99")
            __check((nums[0] + nums[1] + nums[2]).toString(), "139")
        }
