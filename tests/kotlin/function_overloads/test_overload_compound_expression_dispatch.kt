// vybe-test: kotlin/function_overloads/test_overload_compound_expression_dispatch
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun format(v: Int, tag: String = "i"): String = tag + v
        fun format(v: String, tag: String = "s"): String = tag + v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format(4)).toString(), "i4")
            __check((format("x", "#")).toString(), "#x")
        }
