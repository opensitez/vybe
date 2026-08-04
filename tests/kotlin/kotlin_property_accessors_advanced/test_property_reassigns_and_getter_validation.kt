// vybe-test: kotlin/kotlin_property_accessors_advanced/test_property_reassigns_and_getter_validation
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class RangeValue {
            private var _value = 0
            var value: Int
                get() = _value
                set(v) { _value = if (v > 10) 10 else v }
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
            val r = RangeValue()
            r.value = 15
            __p((r.value).toString())
        
__check("10")
}
