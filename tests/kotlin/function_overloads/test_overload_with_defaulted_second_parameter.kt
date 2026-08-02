// vybe-test: kotlin/function_overloads/test_overload_with_defaulted_second_parameter
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun ping(v: Int): String = "solo"
        fun ping(v: Int, label: String = "ok"): String = v.toString() + label
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((ping(1)).toString(), "solo")
            __check((ping(1, "x")).toString(), "1x")
        }
