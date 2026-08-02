// vybe-test: kotlin/spread_arguments/test_spread_with_double_array
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun sum(a: Double, b: Double, c: Double): Double = a + b + c
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val arr = doubleArrayOf(1.0, 2.0)
            __check((sum(3.0, *arr)).toString(), "6.0")
        }
