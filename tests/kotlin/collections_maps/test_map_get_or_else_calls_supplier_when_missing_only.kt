// vybe-test: kotlin/collections_maps/test_map_get_or_else_calls_supplier_when_missing_only
// origin: languages/kotlin/tests/kotlin/test_collections_maps.rs

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
            val scores = mapOf("a" to 1)
            var asked = 0
            val miss = scores.getOrElse("b") {
                asked += 1
                99
            }
            val hit = scores.getOrElse("a") {
                asked += 1
                88
            }
            __p((miss).toString())
            __p((hit).toString())
            __p((asked).toString())
        
__check("99\n1\n1")
}
