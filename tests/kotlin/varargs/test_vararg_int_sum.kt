// vybe-test: kotlin/varargs/test_vararg_int_sum
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun sumAll(vararg values: Int): Int = values.sum()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sumAll(1, 2, 3, 4)).toString(), "10")
        }
