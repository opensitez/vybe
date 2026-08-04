// vybe-test: kotlin/smart_casts/test_type_test_on_inherited_class_chain
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

open class Node
        open class Container : Node()
        class Boxed : Container()

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
            val value: Node = Boxed()
            __p((value is Node).toString())
            __p((value is Container).toString())
            __p((value is Boxed).toString())
        
__check("true\ntrue\ntrue")
}
