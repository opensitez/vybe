// vybe-test: kotlin/function_overloads/test_overload_longer_parameter_list
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun merge(a: Int): String = "a"
        fun merge(a: Int, b: Int): String = "ab"
        fun merge(a: Int, b: Int, c: Int): String = "abc"
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((merge(1)).toString(), "a")
            __check((merge(1, 2)).toString(), "ab")
            __check((merge(1, 2, 3)).toString(), "abc")
        }
