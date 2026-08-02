// vybe-test: kotlin/collections/test_array_singleton_and_size_and_index
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = arrayOf(99)
            __check((nums.size).toString(), "1")
            __check((nums[0]).toString(), "99")
            __check((nums.indices).toString(), "0..0")
        }
