// vybe-test: kotlin/interfaces/test_interface_generic_method_bound_by_implementer
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Formatter {
            fun <T : Number> format(value: T): String
        }

        class IntFormatter : Formatter {
            override fun <T : Number> format(value: T): String = "n:" + value.toInt().toString()
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
            val f: Formatter = IntFormatter()
            __p((f.format(12)).toString())
            __p((f.format(12.4)).toString())
        
__check("n:12\nn:12")
}
