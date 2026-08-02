// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_boolean_toggle
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var on = false
            val toggle = { on = !on }
            toggle()
            toggle()
            __check((on).toString(), "false")
        }
