// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_nested_defaults
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Inner(val id: Int)
        data class Outer(val inner: Inner, val label: String)

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
            val a = Outer(Inner(1), "a")
            val b = a.copy(inner = Inner(2))
            __p((a.inner.id).toString())
            __p((b.inner.id).toString())
            __p((b.label).toString())
        
__check("1\n2\na")
}
