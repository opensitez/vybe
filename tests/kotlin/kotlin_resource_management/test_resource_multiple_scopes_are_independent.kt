// vybe-test: kotlin/kotlin_resource_management/test_resource_multiple_scopes_are_independent
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Flag : AutoCloseable {
            var closeCount = 0
            override fun close() { closeCount += 1 }
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
            val first = Flag()
            val second = Flag()
            first.use { }
            second.use { }
            __p((first.closeCount).toString())
            __p((second.closeCount).toString())
        
__check("1\n1")
}
