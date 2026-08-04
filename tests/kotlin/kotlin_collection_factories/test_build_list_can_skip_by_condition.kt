// vybe-test: kotlin/kotlin_collection_factories/test_build_list_can_skip_by_condition
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

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
            val values = buildList {
                for (v in 1..6) {
                    if (v % 2 == 0) add(v)
                }
            }
            __p((values.joinToString(",")).toString())
        
__check("2,4,6")
}
