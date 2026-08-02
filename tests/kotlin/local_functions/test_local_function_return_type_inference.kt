// vybe-test: kotlin/local_functions/test_local_function_return_type_inference
// origin: languages/kotlin/tests/kotlin/test_local_functions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            fun add(a: Int, b: Int) = a + b
            val value = add(3, 4)
            __check((value).toString(), "7")
        }
