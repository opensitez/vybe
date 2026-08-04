// vybe-test: kotlin/java_util_collections/test_java_collections_min_with_comparator
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

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
            val values = java.util.ArrayList<String>(listOf("aa", "z", "bbb", "c"))
            val shortest = java.util.Collections.min(values, compareBy<String> { it.length })
            val longest = java.util.Collections.max(values, compareBy<String> { it.length })
            __p((shortest).toString())
            __p((longest).toString())
        
__check("z\nbbb")
}
