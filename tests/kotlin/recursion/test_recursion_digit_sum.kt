// vybe-test: kotlin/recursion/test_recursion_digit_sum
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun digitSum(v: Int): Int = if (v == 0) 0 else (v % 10) + digitSum(v / 10)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((digitSum(1234)).toString(), "10")
        }
