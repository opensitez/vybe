// vybe-test: kotlin/sealed_types/test_nested_sealed_hierarchy_with_data
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Command {
            data class Print(val value: String) : Command()
            data class Count(val value: Int) : Command()
        }

        fun execute(command: Command): String {
            return when (command) {
                is Command.Print -> command.value
                is Command.Count -> "count=" + command.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((execute(Command.Print("x"))).toString(), "x")
            __check((execute(Command.Count(4))).toString(), "count=4")
        }
