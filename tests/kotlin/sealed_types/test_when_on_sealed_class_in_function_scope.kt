// vybe-test: kotlin/sealed_types/test_when_on_sealed_class_in_function_scope
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Command
        class Start : Command()
        class Stop : Command()

        fun describe(command: Command): String {
            return when (command) {
                is Start -> "start"
                is Stop -> "stop"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val command: Command = if (true) Start() else Stop()
            __check((describe(command)).toString(), "start")
        }
