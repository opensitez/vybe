// vybe-test: kotlin/kotlin_uuid/test_uuid_timestamped_time_ordering_for_v1_or_not
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
            val v1 = java.util.UUID.fromString("f81d4fae-7dec-11d0-a765-00a0c91e6bf6")
            val v4 = java.util.UUID.randomUUID()
            __p((v1.version()).toString())
            __p((v4.version()).toString())
            __p((v1 != v4).toString())
        
__check("1\n4\ntrue")
}
