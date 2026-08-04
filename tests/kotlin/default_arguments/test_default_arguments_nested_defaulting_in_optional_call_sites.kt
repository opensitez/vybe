// vybe-test: kotlin/default_arguments/test_default_arguments_nested_defaulting_in_optional_call_sites
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun format(a: String, b: String = "B", c: String = "C"): String = a + b + c
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
            __p((format("A")).toString())
            __p((format("A", c = "X")).toString())
            __p((format("A", "Y", "Z")).toString())
        
__check("ABC\nAXX\nAYZ")
}
