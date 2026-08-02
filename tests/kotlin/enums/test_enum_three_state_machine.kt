// vybe-test: kotlin/enums/test_enum_three_state_machine
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class State { START, PROCESS, END }
fun status(s: State): Int { return when (s) { State.START -> 0
State.PROCESS -> 1
State.END -> 2 } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((status(State.PROCESS)).toString(), "1") }
