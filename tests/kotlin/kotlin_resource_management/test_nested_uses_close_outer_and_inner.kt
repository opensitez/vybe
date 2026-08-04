// vybe-test: kotlin/kotlin_resource_management/test_nested_uses_close_outer_and_inner
// origin: languages/kotlin/tests/kotlin/test_kotlin_resource_management.rs

class Token(val name: String) : AutoCloseable {
            var calls: Int = 0
            override fun close() {
                calls += 1
            }
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
            val a = Token("a")
            val b = Token("b")
            a.use {
                b.use {
                    __p((a.calls + b.calls).toString())
                }
            }
            __p((a.calls).toString())
            __p((b.calls).toString())
        
__check("0\n1\n1")
}
