// vybe-test: kotlin/named_arguments/test_named_arguments_with_nested_function_call
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun outer(left: Int, right: Int): Int = left + right
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun inner(a: Int, b: Int, c: Int): Int = a + b + c
            __check((outer(3, right = 4)).toString(), "7")
            __check((inner(a = 1, b = 2, c = 3)).toString(), "6")
        }
