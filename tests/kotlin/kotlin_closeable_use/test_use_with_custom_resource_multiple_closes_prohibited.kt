// vybe-test: kotlin/kotlin_closeable_use/test_use_with_custom_resource_multiple_closes_prohibited
// origin: languages/kotlin/tests/kotlin/test_kotlin_closeable_use.rs

import java.io.Closeable

        class Counted : Closeable {
            var closeCount = 0
            override fun close() { closeCount++ }
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
            val tracked = Counted()
            tracked.use {
                __p((tracked.closeCount).toString())
            }
            __p((tracked.closeCount).toString())
            try {
                tracked.close()
                __p(("extra").toString())
            } catch (e: Exception) {
                __p(("err").toString())
            }
            __p((tracked.closeCount).toString())
        
__check("0\n1\nextra\n2")
}
