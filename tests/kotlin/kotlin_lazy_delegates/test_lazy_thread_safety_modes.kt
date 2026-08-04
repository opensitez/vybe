// vybe-test: kotlin/kotlin_lazy_delegates/test_lazy_thread_safety_modes
// origin: languages/kotlin/tests/kotlin/test_kotlin_lazy_delegates.rs

import kotlin.LazyThreadSafetyMode
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
            var count = 0
            val a by lazy(LazyThreadSafetyMode.NONE) {
                count += 1
                "a"
            }
            val b by lazy(LazyThreadSafetyMode.SYNCHRONIZED) {
                count += 10
                "b"
            }
            __p((a).toString())
            __p((a).toString())
            __p((b).toString())
            __p((b).toString())
            __p((count).toString())
        
__check("a\na\nb\nb\n11")
}
