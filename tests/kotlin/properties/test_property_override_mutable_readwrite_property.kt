// vybe-test: kotlin/properties/test_property_override_mutable_readwrite_property
// origin: languages/kotlin/tests/kotlin/test_properties.rs

interface CounterLike {
            var count: Int
        }

        class Stateful : CounterLike {
            private var raw = 1
            override var count: Int
                get() = raw
                set(next) { raw = next + 1 }
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
            val c: CounterLike = Stateful()
            c.count = 2
            __p((c.count).toString())
        
__check("3")
}
