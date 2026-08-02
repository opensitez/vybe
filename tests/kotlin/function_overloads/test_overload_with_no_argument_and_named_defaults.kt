// vybe-test: kotlin/function_overloads/test_overload_with_no_argument_and_named_defaults
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun flag(v: Int = 1): String = "n" + v
        fun flag(v: String = "s"): String = "s" + v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((flag()).toString(), "n1")
            __check((flag(2)).toString(), "n2")
            __check((flag(v = "x")).toString(), "sx")
        }
