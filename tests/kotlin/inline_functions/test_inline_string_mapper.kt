// vybe-test: kotlin/inline_functions/test_inline_string_mapper
// origin: languages/kotlin/tests/kotlin/test_inline_functions.rs

inline fun format(prefix: String, value: Int, fmt: (Int) -> String): String {
            return prefix + fmt(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format("v=", 4) { v -> v.toString() }).toString(), "v=4")
        }
