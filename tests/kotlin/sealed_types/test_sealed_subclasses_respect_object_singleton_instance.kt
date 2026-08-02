// vybe-test: kotlin/sealed_types/test_sealed_subclasses_respect_object_singleton_instance
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class State {
            object Active : State()
            class Paused(val count: Int) : State()
        }

        fun render(state: State): String {
            return when (state) {
                is State.Active -> "active"
                is State.Paused -> "paused-" + state.count.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((render(State.Active)).toString(), "active")
            __check((render(State.Paused(4))).toString(), "paused-4")
            __check((render(State.Active)).toString(), "active")
        }
