// vybe-test: kotlin/recursion/test_recursion_binary_grow
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun seq(n: Int): Int = if (n <= 1) 1 else seq(n - 1) + seq(n - 2)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((seq(7)).toString(), "13")
        }
