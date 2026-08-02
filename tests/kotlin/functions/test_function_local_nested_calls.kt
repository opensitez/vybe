// vybe-test: kotlin/functions/test_function_local_nested_calls
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun base(x: Int): Int { return x * 2 }
            fun nested(x: Int): Int { return base(x) + 1 }
            __check((nested(3)).toString(), "7")
        }
