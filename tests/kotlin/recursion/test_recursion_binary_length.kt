// vybe-test: kotlin/recursion/test_recursion_binary_length
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun binLen(n: Int): Int = if (n < 2) 1 else 1 + binLen(n / 2)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((binLen(8)).toString(), "4")
            __check((binLen(1)).toString(), "1")
        }
