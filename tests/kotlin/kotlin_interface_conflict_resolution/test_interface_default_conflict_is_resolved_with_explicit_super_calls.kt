// vybe-test: kotlin/kotlin_interface_conflict_resolution/test_interface_default_conflict_is_resolved_with_explicit_super_calls
// origin: languages/kotlin/tests/kotlin/test_kotlin_interface_conflict_resolution.rs

interface First {
            fun origin(): String = "first"
        }

        interface Second {
            fun origin(): String = "second"
        }

        class Composite : First, Second {
            override fun origin(): String = super<First>.origin() + "/" + super<Second>.origin()
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
            __p((Composite().origin()).toString())
        
__check("first/second")
}
