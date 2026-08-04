// vybe-test: kotlin/kotlin_destructuring_maps/test_entry_to_map_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

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
            val map = mapOf("a" to 1, "bb" to 2)
            val out = map
                .toList()
                .associate { (k, v) -> Pair(k + v, v + 1) }
            __p((out["a1"]).toString())
            __p((out["bb2"]).toString())
        
__check("2\n3")
}
