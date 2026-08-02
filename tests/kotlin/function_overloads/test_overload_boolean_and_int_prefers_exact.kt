// vybe-test: kotlin/function_overloads/test_overload_boolean_and_int_prefers_exact
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun value(v: Int): String = "num"
        fun value(v: Boolean): String = "bool"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((value(1)).toString(), "num")
            __check((value(true)).toString(), "bool")
        }
