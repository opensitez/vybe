// vybe-test: kotlin/function_overloads/test_overload_nested_call_resolution
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun call(v: Int): String = "i" + v
        fun call(v: Int, t: String): String = "it" + t
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((call(7)).toString(), "i7")
            __check((call(7, "x")).toString(), "itx")
        }
