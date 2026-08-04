// vybe-test: kotlin/generic_constraints/test_generic_constraints_restrict_to_nullable_char
// origin: languages/kotlin/tests/kotlin/test_generic_constraints.rs

fun <T : CharSequence> safeFirst(v: T?): String = if (v == null || v.isEmpty()) "-" else v[0].toString()
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
            __p((safeFirst("abc")).toString())
            __p((safeFirst(null)).toString())
        
__check("a\n-")
}
