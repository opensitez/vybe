// vybe-test: kotlin/recursion/test_recursion_string_split_depth
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun words(s: String): Int = if (s.isEmpty()) 0 else 1 + words(s.substring(1).trimStart())
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((words("x")).toString(), "1")
        }
