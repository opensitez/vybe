// vybe-test: kotlin/kotlin_closure_capture_mutability/test_capture_data_class_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_closure_capture_mutability.rs

data class State(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val state = State(1)
            val bump = { state.value++ }
            bump()
            bump()
            __check((state.value).toString(), "3")
        }
