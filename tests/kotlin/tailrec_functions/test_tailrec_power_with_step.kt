// vybe-test: kotlin/tailrec_functions/test_tailrec_power_with_step
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun power(base: Int, exp: Int, acc: Int = 1): Int {
            return if (exp == 0) acc else power(base, exp - 1, acc * base)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((power(3, 4)).toString(), "81")
        }
