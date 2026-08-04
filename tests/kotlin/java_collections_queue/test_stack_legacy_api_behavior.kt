// vybe-test: kotlin/java_collections_queue/test_stack_legacy_api_behavior
// origin: languages/kotlin/tests/kotlin/test_java_collections_queue.rs

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
            val stack = java.util.Stack<Int>()
            stack.push(1)
            stack.push(2)
            __p((stack.peek()).toString())
            __p((stack.pop()).toString())
            __p((stack.peek()).toString())
            __p((stack.empty()).toString())
            __p((stack.size).toString())
        
__check("2\n2\n1\nfalse\n1")
}
