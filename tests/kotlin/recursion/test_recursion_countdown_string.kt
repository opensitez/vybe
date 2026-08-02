// vybe-test: kotlin/recursion/test_recursion_countdown_string
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun dump(n: Int): String = if (n <= 0) "0" else n.toString() + "," + dump(n - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((dump(3)).toString(), "3,2,1,0")
        }
