// vybe-test: kotlin/kotlin_uuid/test_uuid_random_uuid_format_shape
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
            val id = java.util.UUID.randomUUID()
            val text = id.toString()
            __p((text.length).toString())
            __p((text[8] == '-').toString())
            __p((text[13] == '-').toString())
            __p((text[18] == '-').toString())
            __p((text[23] == '-').toString())
        
__check("36\ntrue\ntrue\ntrue\ntrue")
}
