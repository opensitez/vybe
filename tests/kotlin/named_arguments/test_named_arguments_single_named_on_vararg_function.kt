// vybe-test: kotlin/named_arguments/test_named_arguments_single_named_on_vararg_function
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun join(prefix: String, vararg values: String, sep: String = ","): String {
            return prefix + values.joinToString(sep)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((join(prefix = "x", values = arrayOf("a", "b"), sep = ":")).toString(), "x:a:b")
            __check((join("x", "1", "2", sep = ";")).toString(), "x1;2")
        }
