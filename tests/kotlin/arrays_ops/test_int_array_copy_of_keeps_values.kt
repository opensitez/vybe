// vybe-test: kotlin/arrays_ops/test_int_array_copy_of_keeps_values
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = intArrayOf(1, 2, 3)
            val copy = base.copyOf()
            copy[0] = 10
            __check((base[0]).toString(), "1")
            __check((copy[0]).toString(), "10")
            __check((base.size).toString(), "3")
        }
