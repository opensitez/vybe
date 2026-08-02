// vybe-test: kotlin/arrays_ops/test_int_array_factory_lambda
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = IntArray(4) { idx -> idx * idx }
            __check((nums.joinToString(",")).toString(), "0,1,4,9")
        }
