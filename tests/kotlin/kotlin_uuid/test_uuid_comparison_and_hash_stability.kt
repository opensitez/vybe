// vybe-test: kotlin/kotlin_uuid/test_uuid_comparison_and_hash_stability
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

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
            val a = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            val b = java.util.UUID.fromString("123e4567-e89b-12d3-a456-426614174000")
            __p((a == b).toString())
            __p((a.hashCode() == b.hashCode()).toString())
        
__check("true\ntrue")
}
