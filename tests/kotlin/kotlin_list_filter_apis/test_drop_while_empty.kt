// vybe-test: kotlin/kotlin_list_filter_apis/test_drop_while_empty
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_filter_apis.rs

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
            val nums = listOf(1, 2, 3)
            __p((nums.dropWhile { it < 0 }.joinToString(",")).toString())
            __p((nums.takeWhile { false }.joinToString(",")).toString())
            __p((nums.takeWhile { true }.size).toString())
        
__check("1,2,3\n\n3")
}
