// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_function_reference
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun make(prefix: String): (Int) -> String {
            return { v -> prefix + v.toString() }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = make("x")
            __check((f(7)).toString(), "x7")
        }
