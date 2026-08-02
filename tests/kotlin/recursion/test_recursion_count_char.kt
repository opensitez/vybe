// vybe-test: kotlin/recursion/test_recursion_count_char
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun countChar(s: String, target: Char): Int {
            if (s.isEmpty()) return 0
            val head = if (s[0] == target) 1 else 0
            return head + countChar(s.substring(1), target)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((countChar("abca", 'a')).toString(), "2")
        }
