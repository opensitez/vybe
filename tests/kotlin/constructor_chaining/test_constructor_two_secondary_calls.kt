// vybe-test: kotlin/constructor_chaining/test_constructor_two_secondary_calls
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Layer {
            val a: Int
            val b: Int
            constructor(a: Int) { this.a = a
this.b = a }
            constructor(a: Int, b: Int) : this(a + b)
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
            val l = Layer(2, 3)
            __p((l.a).toString())
            __p((l.b).toString())
        
__check("5\n5")
}
