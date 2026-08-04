// vybe-test: kotlin/data_classes/test_data_class_in_map_for_hash_lookup_after_copy
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Key(val id: Int, val label: String)

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
            val original = Key(1, "a")
            val map = mutableMapOf(original to "first")
            val copy = original.copy()
            original.label = "b"
            __p((map[original] == null).toString())
            __p((map[copy]).toString())
        
__check("true\nfirst")
}
