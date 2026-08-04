// vybe-test: kotlin/data_class_copying/test_data_class_copy_replaces_multiple_properties
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Box(val id: Int, val tag: String, val active: Boolean)
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
            val a = Box(1, "x", false)
            val b = a.copy(id = 2, active = true)
            __p((a.id).toString())
            __p((b.id).toString())
            __p((b.tag).toString())
            __p((b.active).toString())
        
__check("1\n2\nx\ntrue")
}
