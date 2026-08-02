// vybe-test: kotlin/functions/test_function_nested_capture_uses_outer_state
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun makeAdder(base: Int): (Int) -> Int {
            return { delta ->
                base + delta
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val addTen = makeAdder(10)
            __check((addTen(3)).toString(), "13")
        }
