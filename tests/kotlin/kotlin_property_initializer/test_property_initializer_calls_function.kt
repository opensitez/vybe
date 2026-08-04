// vybe-test: kotlin/kotlin_property_initializer/test_property_initializer_calls_function
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_initializer.rs

var counter = 0

        fun next(): Int {
            counter = counter + 1
            return counter
        }

        class Node {
            val a: Int = next()
            val b: Int = next()
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
            val n = Node()
            __p((n.a).toString())
            __p((n.b).toString())
            __p((counter).toString())
        
__check("1\n2\n2")
}
