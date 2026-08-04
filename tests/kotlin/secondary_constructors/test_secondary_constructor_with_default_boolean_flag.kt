// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_with_default_boolean_flag
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Marker {
            val active: Boolean
            val label: String

            constructor(label: String) {
                this.label = label
                this.active = false
            }

            constructor(label: String, active: Boolean) : this(label) {
                if (active) this.label = label + "!"
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
            val a = Marker("x")
            val b = Marker("x", true)
            __p((a.active).toString())
            __p((a.label).toString())
            __p((b.label).toString())
        
__check("false\nx\nx!")
}
