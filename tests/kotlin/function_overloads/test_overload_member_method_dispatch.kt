// vybe-test: kotlin/function_overloads/test_overload_member_method_dispatch
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

class Solver {
            fun eval(v: Int): Int = v + 1
            fun eval(v: String): String = v + "!"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = Solver()
            __check((s.eval(3)).toString(), "4")
            __check((s.eval("x")).toString(), "x!")
        }
