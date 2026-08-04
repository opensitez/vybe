// vybe-test: kotlin/data_class_destructuring/test_destructure_with_default_values_is_explicit_constructor
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Node(val value: Int = 4, val name: String = "x")

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
            val one = Node(name = "n")
            val two = Node(3, "m")
            val (a, b) = one
            val (c, d) = two
            __p((a).toString())
            __p((b).toString())
            __p((c).toString())
            __p((d).toString())
        
__check("4\nn\n3\nm")
}
