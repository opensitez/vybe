// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_var_with_default_argument
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var text = "a"
            val append = { suffix: String -> text += suffix }
            append("b")
            append("c")
            __check((text).toString(), "abc")
        }
