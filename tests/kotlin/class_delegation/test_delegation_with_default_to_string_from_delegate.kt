// vybe-test: kotlin/class_delegation/test_delegation_with_default_to_string_from_delegate
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Named { fun title(): String }

        class NamedImpl : Named {
            override fun title() = "named"
            override fun toString() = "impl"
        }

        class NamedProxy(delegate: Named) : Named by delegate

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
            val value = NamedProxy(NamedImpl())
            __p((value.title()).toString())
            __p((value.toString()).toString())
        
__check("named\nimpl")
}
