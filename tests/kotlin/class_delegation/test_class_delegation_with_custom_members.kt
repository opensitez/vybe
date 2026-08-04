// vybe-test: kotlin/class_delegation/test_class_delegation_with_custom_members
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Counter {
            fun value(): Int
        }

        class NumberCounter(private val v: Int) : Counter {
            override fun value() = v
        }

        class OffsetCounter(base: Counter) : Counter by base {
            fun id() = "offset"
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
            val c = OffsetCounter(NumberCounter(3))
            __p((c.id()).toString())
            __p((c.value()).toString())
        
__check("offset\n3")
}
