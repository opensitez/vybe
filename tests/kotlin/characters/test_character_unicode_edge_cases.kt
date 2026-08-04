// vybe-test: kotlin/characters/test_character_unicode_edge_cases
// origin: languages/kotlin/tests/kotlin/test_characters.rs

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
            val space = '\u0020'
            val tab = '\u0009'
            val euro = '€'
            __p((space.isWhitespace()).toString())
            __p((tab.isWhitespace()).toString())
            __p((euro.isLetterOrDigit()).toString())
            __p((euro.isDefined()).toString())
            __p((euro.code).toString())
        
__check("true\ntrue\nfalse\ntrue\n8364")
}
