// vybe-test: kotlin/recursion/test_recursion_string_builder
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

fun repeatChar(ch: Char, count: Int): String {
            return if (count <= 0) "" else ch + repeatChar(ch, count - 1)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((repeatChar('x', 3)).toString(), "xxx")
        }
