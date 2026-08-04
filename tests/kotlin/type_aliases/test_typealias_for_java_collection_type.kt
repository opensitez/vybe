// vybe-test: kotlin/type_aliases/test_typealias_for_java_collection_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias JavaMap = java.util.LinkedHashMap<String, Int>

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
            val counts: JavaMap = java.util.LinkedHashMap<String, Int>()
            counts["a"] = 1
            counts["b"] = 2
            counts.put("a", 3)
            __p((counts["a"]).toString())
            __p((counts.size).toString())
        
__check("3\n2")
}
