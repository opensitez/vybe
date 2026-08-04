// vybe-test: kotlin/collections_maps/test_map_map_keys_and_values_transform
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
            val input = mapOf("a" to 1, "b" to 2)
            val keys = input.mapKeys { it.key.uppercase() }
            val values = input.mapValues { it.value * 10 }
            __p((keys["A"]).toString())
            __p((values["b"]).toString())
            __p((keys.size).toString())
        
__check("1\n20\n2")
}
