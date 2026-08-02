// vybe-test: kotlin/varargs/test_vararg_extension_function
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun String.wrapAll(vararg values: String): String = values.joinToString(this, prefix = "<", postfix = ">")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((",".wrapAll("x", "y", "z")).toString(), "<x,y,z>")
        }
