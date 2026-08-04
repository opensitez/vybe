// vybe-test: kotlin/collections_iterables/test_mutable_list_operations_modify_size_and_order
// origin: languages/kotlin/tests/kotlin/test_collections_iterables.rs

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
            val nums = mutableListOf(1, 2, 3)
            nums.add(4)
            nums.removeAt(1)
            nums[0] = 8
            nums.remove(3)
            __p((nums.joinToString(",")).toString())
            __p((nums.size).toString())
        
__check("8,4\n2")
}
