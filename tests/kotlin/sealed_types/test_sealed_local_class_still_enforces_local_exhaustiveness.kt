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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((decide(State.Ok())).toString())
            __p((decide(State.Fail())).toString())
            __p((decide(State.Ignore())).toString())
        
__check("ok\nfail\nignore")
}
