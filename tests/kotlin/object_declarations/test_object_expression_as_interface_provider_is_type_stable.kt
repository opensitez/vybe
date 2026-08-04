// vybe-test: kotlin/object_declarations/test_object_expression_as_interface_provider_is_type_stable
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Printer {
            fun __pr(""): String
        }

        fun makePrinter(prefix: String): Printer {
            return object : Printer {
                override fun __pr(""): String = prefix + "!"
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
            val first = makePrinter("a")
            val second = makePrinter("b")
            __p((first.print()).toString())
            __p((second.print()).toString())
        
__check("a!\nb!")
}
