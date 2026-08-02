// vybe-test: kotlin/varargs/test_vararg_used_alongside_single_list_argument
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun append(base: String, separator: String, vararg values: String): String {
            return base + values.joinToString(separator)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((append("v", ".", "a", "b", "c")).toString(), "v.a.b.c")
        }
