// vybe-test: kotlin/equality_hashcode/test_data_class_copy_changes_only_targeted_fields
// origin: languages/kotlin/tests/kotlin/test_equality_hashcode.rs

data class Pair(val left: Int, val right: String)

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
            val source = Pair(1, "a")
            val updated = source.copy(left = 3)
            __p((source.left).toString())
            __p((updated.left).toString())
            __p((updated.right).toString())
            __p((source.right).toString())
        
__check("1\n3\na\na")
}
