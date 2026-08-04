// vybe-test: kotlin/kotlin_string_line_ops/test_string_replace_and_split_ops
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_line_ops.rs

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
            __p(("aa-bb-cc".replaceFirst("-", "/")).toString())
            __p(("aa-bb-cc".replace("-", "/")).toString())
            val out = "a,b,c".split(",")
            __p((out.size).toString())
            __p((out[1]).toString())
            __p(("a,b,c".split(",", limit = 2).size).toString())
        
__check("aa/bb-cc\naa/bb/cc\n3\nb\n2")
}
