// vybe-test: kotlin/class_delegation/test_delegate_when_base_state_changes
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface MutableCounter { var value: Int }

        class Counter(var value: Int) : MutableCounter

        class Proxy(delegate: MutableCounter) : MutableCounter by delegate

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
            val c = Counter(1)
            val p = Proxy(c)
            p.value = p.value + 2
            __p((p.value).toString())
            __p((c.value).toString())
        
__check("3\n3")
}
