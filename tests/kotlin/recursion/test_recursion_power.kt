// vybe-test: kotlin/recursion/test_recursion_power
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun pow(base: Int, exp: Int): Int = if (exp == 0) 1 else base * pow(base, exp - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pow(2, 4)).toString(), "16")
        }
