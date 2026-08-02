// vybe-test: kotlin/recursion/test_recursion_path_sum
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun path(n: Int, acc: Int): Int = if (n <= 0) acc else path(n - 1, acc + n)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((path(3, 0)).toString(), "6")
        }
