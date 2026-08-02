// vybe-test: kotlin/recursion/test_recursion_head_sum
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun headSum(values: List<Int>): Int = if (values.isEmpty()) 0 else values[0] + headSum(values.drop(1))
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((headSum(listOf(2, 4, 6))).toString(), "12")
        }
