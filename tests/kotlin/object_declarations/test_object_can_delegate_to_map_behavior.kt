// vybe-test: kotlin/object_declarations/test_object_can_delegate_to_map_behavior
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Cache : Map<String, Int> by mapOf("a" to 1, "b" to 2) {
            val keysText = keys.joinToString("-")
        }

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
            __p((Cache["a"]).toString())
            __p((Cache.keysText).toString())
            __p((Cache.size).toString())
        
__check("1\na-b\n2")
}
