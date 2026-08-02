// vybe-test: kotlin/recursion/test_recursion_repeat_list
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun repeat(value: String, count: Int): String = if (count <= 0) "" else value + repeat(value, count - 1)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((repeat("a", 2)).toString(), "aa")
        }
