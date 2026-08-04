// vybe-test: kotlin/data_classes/test_data_class_in_set_with_mutated_var_field_can_change_lookup
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Packet(var id: Int, val payload: String)

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
            val one = Packet(1, "x")
            val set = mutableSetOf(one)
            __p((set.contains(Packet(1, "x"))).toString())
            one.id = 2
            __p((set.contains(Packet(2, "x"))).toString())
            __p((set.contains(Packet(1, "x"))).toString())
        
__check("true\nfalse\nfalse")
}
