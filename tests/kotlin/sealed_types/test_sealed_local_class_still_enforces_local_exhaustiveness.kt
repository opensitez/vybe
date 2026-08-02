// vybe-test: kotlin/sealed_types/test_sealed_local_class_still_enforces_local_exhaustiveness
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

fun decide(state: State): String {
            return when (state) {
                is State.Ok -> "ok"
                is State.Fail -> "fail"
                is State.Ignore -> "ignore"
            }
        }

        sealed class State {
            class Ok : State()
            class Fail : State()
            class Ignore : State()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((decide(State.Ok())).toString(), "ok")
            __check((decide(State.Fail())).toString(), "fail")
            __check((decide(State.Ignore())).toString(), "ignore")
        }
