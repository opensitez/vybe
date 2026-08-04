// vybe-test: kotlin/map_lookup_projection/test_map_get_or_put_build_lists
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

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
            val buckets = linkedMapOf<String, MutableList<Int>>()
            buckets.getOrPut("a") { mutableListOf() }.add(1)
            buckets.getOrPut("a") { mutableListOf() }.add(2)
            buckets.getOrPut("b") { mutableListOf() }.add(9)
            __p((buckets["a"]!!.joinToString(",")).toString())
            __p((buckets.size).toString())
        
__check("1,2\n2")
}
