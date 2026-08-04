// vybe-test: kotlin/property_accessors/test_property_backed_by_computed_expression
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

class Box {
            var a: Int = 2
            var b: Int = 3
            val sum: Int
                get() = a + b
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
            val b = Box()
            __p((b.sum).toString())
            b.a = 5
            __p((b.sum).toString())
        
__check("5\n8")
}
