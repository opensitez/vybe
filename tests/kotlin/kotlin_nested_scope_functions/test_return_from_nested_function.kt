// vybe-test: kotlin/kotlin_nested_scope_functions/test_return_from_nested_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_nested_scope_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun compute(x: Int): Int {
                fun square(v: Int) = v * v
                return square(x) + 1
            }
            __check((compute(4)).toString(), "17")
        }
