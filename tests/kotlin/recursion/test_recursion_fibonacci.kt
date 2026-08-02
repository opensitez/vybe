// vybe-test: kotlin/recursion/test_recursion_fibonacci
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun fib(n: Int): Int = if (n <= 1) n else fib(n - 1) + fib(n - 2)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((fib(6)).toString(), "8")
        }
