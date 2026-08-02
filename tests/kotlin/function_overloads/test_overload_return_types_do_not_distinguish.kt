// vybe-test: kotlin/function_overloads/test_overload_return_types_do_not_distinguish
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun value(v: Int): Int = v
        fun value(v: Int, b: Int): Int = v + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((value(2)).toString(), "2")
            __check((value(2, 3)).toString(), "5")
        }
