// vybe-test: kotlin/arrays_ops/test_int_array_copy_of_with_new_size
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3)
            val grown = nums.copyOf(5)
            val shrunk = nums.copyOf(2)
            __check((grown.joinToString(",")).toString(), "1,2,3,0,0")
            __check((shrunk.joinToString(",")).toString(), "1,2")
        }
