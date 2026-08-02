// vybe-test: kotlin/arrays_ops/test_int_array_fold_and_reduce_contracts
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            __check((nums.reduce { acc, value -> acc + value }).toString(), "10")
            __check((nums.fold(10) { acc, value -> acc - value }).toString(), "0")
            __check((nums.fold("") { acc, value -> acc + value.toString() }).toString(), "1234")
        }
