// vybe-test: kotlin/kotlin_type_checks/test_is_as_and_as_question_mark_paths
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_checks.rs

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
            val values: List<Any?> = listOf("x", 2, null, 3.1)
            val first = values[0] as String
            val second = values[2] as? String
            __p((first).toString())
            __p((second).toString())
        
__check("x\nnull")
}
