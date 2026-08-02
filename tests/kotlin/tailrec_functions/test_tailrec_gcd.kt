// vybe-test: kotlin/tailrec_functions/test_tailrec_gcd
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun gcd(a: Int, b: Int): Int {
            return if (b == 0) a else gcd(b, a % b)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((gcd(48, 18)).toString(), "6")
        }
