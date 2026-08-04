// vybe-test: kotlin/constructor_chaining/test_constructor_init_read_order
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Bag(val base: Int) {
            val a: Int
            init { a = base + 1 }
            val b = a + 2
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
            val b = Bag(3)
            __p((b.a).toString())
            __p((b.b).toString())
        
__check("4\n6")
}
