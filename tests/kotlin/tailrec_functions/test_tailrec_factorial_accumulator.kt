// vybe-test: kotlin/tailrec_functions/test_tailrec_factorial_accumulator
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun factorial(n: Int, acc: Int = 1): Int {
            return if (n <= 1) acc else factorial(n - 1, n * acc)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((factorial(5)).toString(), "120")
        }
