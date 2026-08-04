// vybe-test: kotlin/kotlin_map_apis/test_map_filter_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_apis.rs

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
            val map = linkedMapOf("a" to 1, "b" to 4, "c" to 2)
            val filtered = map.filterValues { it >= 3 }
            __p((filtered.size).toString())
            __p((filtered["b"]).toString())
            __p((filtered.containsKey("a")).toString())
        
__check("1\n4\nfalse")
}
