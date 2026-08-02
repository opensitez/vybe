// vybe-test: kotlin/function_overloads/test_overload_boolean_vs_unit_not_possible
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun ping(v: Int): String = "i"
        fun ping(v: Boolean, force: Int = 0): String = "b" + force
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ping(1)).toString(), "i")
            __check((ping(true)).toString(), "b0")
            __check((ping(false, 2)).toString(), "b2")
        }
