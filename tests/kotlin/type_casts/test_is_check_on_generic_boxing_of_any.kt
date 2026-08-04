// vybe-test: kotlin/type_casts/test_is_check_on_generic_boxing_of_any
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun isStringList(value: Any): Boolean {
            return value is List<*>
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
            __p((isStringList(listOf("a", "b", "c"))).toString())
            __p((isStringList(10)).toString())
            val maybeList: Any? = null
            __p((maybeList is List<*>).toString())
        
__check("true\nfalse\nfalse")
}
