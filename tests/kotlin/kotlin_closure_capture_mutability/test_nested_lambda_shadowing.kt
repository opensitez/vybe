// vybe-test: kotlin/kotlin_closure_capture_mutability/test_nested_lambda_shadowing
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var value = 1
            val outer = {
                val value = 10
                { value + 1 }
            }
            __check((outer()()).toString(), "11")
            __check((value).toString(), "1")
        }
