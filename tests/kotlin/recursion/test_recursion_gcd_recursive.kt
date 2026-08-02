// vybe-test: kotlin/recursion/test_recursion_gcd_recursive
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun gcd(a: Int, b: Int): Int = if (b == 0) a else gcd(b, a % b)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((gcd(20, 12)).toString(), "4")
        }
