// vybe-test: kotlin/scoping_functions/test_run_supports_returning_receiver_after_mutation_with_lambda
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class State(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = State(1)
            val copy = source.run {
                value = value + 4
                this
            }
            __check((source.value).toString(), "5")
            __check((copy.value).toString(), "5")
            __check((source === copy).toString(), "true")
        }
