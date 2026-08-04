// vybe-test: kotlin/type_casts/test_casting_array_to_primitive_array_projection_is_type_checked
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

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
            val values: Any = arrayOf(5, 6, 7)
            val primitive = values as? IntArray
            val boxed = values as? Array<Int>
            __p((primitive == null).toString())
            __p((boxed != null).toString())
            __p((boxed?.size).toString())
        
__check("true\ntrue\n3")
}
