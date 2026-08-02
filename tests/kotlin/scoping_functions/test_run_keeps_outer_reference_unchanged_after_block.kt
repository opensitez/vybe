// vybe-test: kotlin/scoping_functions/test_run_keeps_outer_reference_unchanged_after_block
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class State(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val state = State(4)
            val block = state.run {
                value = value * 3
                value + 1
            }
            __check((state.value).toString(), "12")
            __check((block).toString(), "13")
        }
