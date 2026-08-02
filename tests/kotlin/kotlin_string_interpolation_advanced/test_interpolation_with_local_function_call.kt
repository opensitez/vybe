// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_with_local_function_call
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun value(a: Int) = a * 2
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x = 3
            __check(("doubled=${'$'}{value(x)}").toString(), "doubled=6")
        }
