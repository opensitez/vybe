// vybe-test: kotlin/inline_functions/test_inline_void_style
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun tap(value: Int, action: (Int) -> Unit): Int {
            action(value)
            return value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var seen = 0
            val out = tap(3) { v -> seen += v }
            __check((out).toString(), "3")
            __check((seen).toString(), "3")
        }
