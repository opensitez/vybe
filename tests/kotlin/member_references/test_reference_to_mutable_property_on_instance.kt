// vybe-test: kotlin/member_references/test_reference_to_mutable_property_on_instance
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Holder {
            var value: Int = 0
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
            val holder = Holder()
            holder.value = 8
            val read = holder::value
            holder.value = 1
            val read2 = Holder::value
            __p((read()).toString())
            __p((read2(holder)).toString())
        
__check("1\n1")
}
