// vybe-test: kotlin/tailrec_functions/test_tailrec_even_or_odd_classifier
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun parity(n: Int): String {
            return if (n == 0) "even" else if (n == 1) "odd" else parity(n - 2)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((parity(9)).toString(), "odd")
        }
