// vybe-test: kotlin/properties/test_property_setter_records_previous_state
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Logarithm {
            private var raw = 0
            var value: Int
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
            val value = Logarithm()
            value.value = 3
            value.value = 7
            __p((value.value).toString())
        
__check("8")
}
