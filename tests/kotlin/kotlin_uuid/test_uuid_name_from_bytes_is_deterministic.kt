// vybe-test: kotlin/kotlin_uuid/test_uuid_name_from_bytes_is_deterministic
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
            val bytes = "kotlin".toByteArray()
            val a = java.util.UUID.nameUUIDFromBytes(bytes)
            val b = java.util.UUID.nameUUIDFromBytes(bytes)
            __p((a).toString())
            __p((a == b).toString())
            __p((a.version()).toString())
        
__check("7f1d0d8e-2f1f-3138-92e7-4b3d8f5ef2d6\ntrue\n3")
}
