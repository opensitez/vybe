// vybe-test: kotlin/function_overloads/test_overload_with_vararg_and_single_arg
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun take(v: String): String = "one"
        fun take(vararg v: String): String = "many:" + v.size
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((take("a")).toString(), "one")
            __check((take("a", "b", "c")).toString(), "many:3")
        }
