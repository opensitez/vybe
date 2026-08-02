// vybe-test: kotlin/varargs/test_vararg_nullable_elements
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun read(values: vararg value: String?): String {
            return values.joinToString(";") { it ?: "nil" }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((read("x", null, "z")).toString(), "x;nil;z")
        }
