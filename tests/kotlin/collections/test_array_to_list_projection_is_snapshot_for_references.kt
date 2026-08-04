// vybe-test: kotlin/collections/test_array_to_list_projection_is_snapshot_for_references
// origin: languages/kotlin/tests/kotlin/test_collections.rs

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
            val nums = IntArray(3) { it + 1 }
            val snapshot = nums.toList()
            nums[0] = 9
            __p((snapshot.joinToString(",")).toString())
            __p((nums.joinToString(",")).toString())
            __p((snapshot[1]).toString())
        
__check("1,2,3\n9,2,3\n2")
}
