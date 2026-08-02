// vybe-test: kotlin/arrays_ops/test_array_take_last_and_drop
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4, 5)
            val tail = nums.takeLast(2)
            val mid = nums.drop(2)
            __check((tail.joinToString(",")).toString(), "4,5")
            __check((mid.joinToString(",")).toString(), "3,4,5")
        }
