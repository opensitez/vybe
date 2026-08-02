// vybe-test: kotlin/kotlin_control_flow_guards/test_if_with_guarded_branching
// origin: languages/kotlin/tests/kotlin/test_kotlin_control_flow_guards.rs

fun classify(v: Int): String = if (v > 0) {
            "positive"
        } else if (v < 0) {
            "negative"
        } else {
            "zero"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(4)).toString(), "positive")
            __check((classify(-1)).toString(), "negative")
            __check((classify(0)).toString(), "zero")
        }
