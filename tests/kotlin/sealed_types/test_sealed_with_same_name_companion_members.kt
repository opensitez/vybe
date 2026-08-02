// vybe-test: kotlin/sealed_types/test_sealed_with_same_name_companion_members
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class State {
            class Active : State()
            class Error : State()

            companion object {
                fun active(): State = Active()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val state = State.active()
            __check((when (state) {
                is State.Active -> 1
                is State.Error -> 0
            }).toString(), "1")
        }
