// vybe-test: kotlin/default_arguments/test_default_arguments_for_local_function
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun make(base: Int = 1, extra: Int = 2): Int = base + extra
            __check((make()).toString(), "3")
            __check((make(5)).toString(), "7")
            __check((make(extra = 10, base = 1)).toString(), "11")
        }
