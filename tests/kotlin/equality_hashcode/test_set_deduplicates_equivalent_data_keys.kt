// vybe-test: kotlin/equality_hashcode/test_set_deduplicates_equivalent_data_keys
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Key(val id: String, val version: Int)

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
            val set = hashSetOf(Key("x", 1), Key("x", 1), Key("x", 2))
            __p((set.size).toString())
            __p((set.contains(Key("x", 2))).toString())
            __p((set.contains(Key("x", 3))).toString())
        
__check("2\ntrue\nfalse")
}
