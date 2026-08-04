// vybe-test: kotlin/strings_regex/test_regex_find_all_with_indexes
// origin: languages/kotlin/tests/kotlin/test_strings_regex.rs

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
            val pattern = Regex("[A-Za-z]")
            val first = pattern.find("A1b2C3")
            __p((first?.value ?: "none").toString())
            val all = pattern.findAll("A1b2C3").toList()
            __p((all[0].range.start).toString())
            __p((all[1].range.start).toString())
            __p((all[2].range.start).toString())
        
__check("A\n0\n2\n4")
}
