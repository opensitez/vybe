// vybe-test: kotlin/kotlin_interface_defaults/test_interface_default_method_is_used_when_not_overridden
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_defaults.rs

interface Logger {
            fun prefix(): String = "log"
            fun format(message: String): String = prefix() + ":" + message
        }

        class DefaultLogger : Logger

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
            __p((DefaultLogger().format("ok")).toString())
        
__check("log:ok")
}
