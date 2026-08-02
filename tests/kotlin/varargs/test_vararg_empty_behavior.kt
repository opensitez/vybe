// vybe-test: kotlin/varargs/test_vararg_empty_behavior
// origin: languages/kotlin/tests/kotlin/test_varargs.rs

fun count(vararg values: Int): Int = values.size

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((count()).toString(), "0")
        }
