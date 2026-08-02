// vybe-test: kotlin/function_overloads/test_overload_by_argument_count
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun token(x: Int): String = "int:" + x
        fun token(x: String, y: String): String = x + y
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((token(4)).toString(), "int:4")
            __check((token("a", "b")).toString(), "ab")
        }
