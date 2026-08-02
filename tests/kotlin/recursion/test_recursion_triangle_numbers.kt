// vybe-test: kotlin/recursion/test_recursion_triangle_numbers
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun triangle(n: Int): Int = if (n <= 0) 0 else n + triangle(n - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((triangle(5)).toString(), "15")
        }
