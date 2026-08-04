// vybe-test: kotlin/nested_classes/test_nested_class_method_dispatch
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Service {
            class State(val ok: Boolean)
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
            val a = Service.State(true)
            val b = Service.State(false)
            val total = (if (a.ok) 1 else 0) + (if (b.ok) 1 else 0)
            __p((total).toString())
        
__check("1")
}
