// vybe-test: kotlin/kotlin_property_backing_state/test_property_getter_derived_field
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_backing_state.rs

class Timer {
            private var total = 0
            var seconds: Int
                get() = total
                set(value) {
                    total = value
                }
            val isZero: Boolean
                get() = total == 0
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
            val t = Timer()
            __p((t.isZero).toString())
            t.seconds = 3
            __p((t.seconds).toString())
            __p((t.isZero).toString())
        
__check("true\n3\nfalse")
}
