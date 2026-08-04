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
            __p((execute(Command.Print("x"))).toString())
            __p((execute(Command.Count(4))).toString())
        
__check("x\ncount=4")
}
