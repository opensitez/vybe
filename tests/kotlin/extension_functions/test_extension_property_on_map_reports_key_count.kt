// vybe-test: kotlin/extension_functions/test_extension_property_on_map_reports_key_count
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

val Map<String, Int>.keyText: String
            get() = keys.joinToString("|")

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
            val values = linkedMapOf("a" to 1, "b" to 2, "c" to 3)
            __p((values.keyText).toString())
            __p((values.keyText).toString())
        
__check("a|b|c\na|b|c")
}
