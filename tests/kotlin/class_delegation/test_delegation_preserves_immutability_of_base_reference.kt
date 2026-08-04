// vybe-test: kotlin/class_delegation/test_delegation_preserves_immutability_of_base_reference
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface View { fun size(): Int }

        class Snapshot(private val items: List<Int>) : View {
            override fun size() = items.size
        }

        class SnapshotWrapper(delegate: View) : View by delegate

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
            val original = Snapshot(listOf(1, 2))
            val wrapped = SnapshotWrapper(original)
            __p((wrapped.size()).toString())
        
__check("2")
}
