// vybe-test: kotlin/arrays_ops/test_array_filter_to_intarray
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5, 6)
            val evens = nums.filter { it % 2 == 0 }.toIntArray()
            __check((evens.joinToString(",")).toString(), "2,4,6")
        }
