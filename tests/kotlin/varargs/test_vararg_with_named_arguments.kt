// vybe-test: kotlin/varargs/test_vararg_with_named_arguments
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun build(prefix: String, suffix: String = "#", vararg values: String): String {
            return prefix + values.joinToString(suffix)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((build("v", ";", "one", "two")).toString(), "vone;two")
        }
