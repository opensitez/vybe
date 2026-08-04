// vybe-test: kotlin/kotlin_atomic_primitives/test_atomic_reference_compare_and_set_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_atomic_primitives.rs

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
            val ref = java.util.concurrent.atomic.AtomicReference(5)
            val ok = ref.compareAndSet(4, 7)
            __p((ok).toString())
            __p((ref.get()).toString())
        
__check("false\n5")
}
