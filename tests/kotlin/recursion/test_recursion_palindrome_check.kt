// vybe-test: kotlin/recursion/test_recursion_palindrome_check
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun isPal(s: String): Boolean {
            if (s.length <= 1) return true
            if (s.first() != s.last()) return false
            return isPal(s.substring(1, s.length - 1))
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isPal("racecar")).toString(), "true")
            __check((isPal("kotlin")).toString(), "false")
        }
