// vybe-test: kotlin/function_overloads/test_overload_nested_call_with_same_name_and_return
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun wrap(v: Int): Int = v + 1
        fun wrap(v: Int, depth: Int): Int = v + depth
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun run(value: Int): Int = wrap(value)
            __check((run(1)).toString(), "2")
            __check((wrap(1, 9)).toString(), "10")
        }
