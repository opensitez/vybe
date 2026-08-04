// vybe-test: kotlin/arrays_ops/test_array_of_nulls_defaults_to_null_and_can_mutate_indices
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

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
            val slots = arrayOfNulls<String>(3)
            val before = slots.joinToString(",") { it ?: "null" }
            slots[1] = "value"
            val after = slots.joinToString(",") { it ?: "null" }
            __p((before).toString())
            __p((after).toString())
            __p((slots.count { it == null }).toString())
        
__check("null,null,null\nnull,value,null\n2")
}
