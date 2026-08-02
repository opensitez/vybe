// vybe-test: kotlin/function_overloads/test_overload_with_default_and_no_default_call
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun mark(v: Int): String = "single"
        fun mark(v: Int, suffix: String = ""): String = "double:" + v + suffix
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((mark(1)).toString(), "double:1")
            __check((mark(2, "ok")).toString(), "double:2ok")
        }
