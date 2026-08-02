// vybe-test: kotlin/varargs/test_vararg_array_list_conversion
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun asList(prefix: String, vararg values: Int): String {
            val list = values.toList().map { it.toString() }.joinToString(prefix = prefix, separator = "-")
            return list
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asList("a", 1, 2, 3)).toString(), "1-2-3")
        }
