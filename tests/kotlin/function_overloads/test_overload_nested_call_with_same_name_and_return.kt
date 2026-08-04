// vybe-test: kotlin/function_overloads/test_overload_nested_call_with_same_name_and_return
// origin: languages/kotlin/tests/kotlin/test_function_overloads.rs

fun wrap(v: Int): Int = v + 1
        fun wrap(v: Int, depth: Int): Int = v + depth
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
            fun run(value: Int): Int = wrap(value)
            __p((run(1)).toString())
            __p((wrap(1, 9)).toString())
        
__check("2\n10")
}
