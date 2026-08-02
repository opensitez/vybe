// vybe-test: kotlin/arrays_ops/test_array_content_hash_code_is_positive
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            val nested = arrayOf(nums)
            __check((nums.contentHashCode() > 0).toString(), "true")
            __check((nested.contentDeepHashCode() > 0).toString(), "true")
        }
