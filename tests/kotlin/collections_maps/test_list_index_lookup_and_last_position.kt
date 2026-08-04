// vybe-test: kotlin/collections_maps/test_list_index_lookup_and_last_position
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
            val values = listOf(5, 6, 7, 6, 8)
            __p((values.indexOf(6)).toString())
            __p((values.lastIndexOf(6)).toString())
            var output = ""
            for (i in values.indices) {
                if (i % 2 == 1) {
                    output += values[i].toString()
                }
            }
            __p((values.size).toString())
            __p((output).toString())
        
__check("1\n3\n4\n68")
}
