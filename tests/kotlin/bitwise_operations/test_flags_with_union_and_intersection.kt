// vybe-test: kotlin/bitwise_operations/test_flags_with_union_and_intersection
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

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
            val canRead = 0b001
            val canWrite = 0b010
            val canExecute = 0b100
            val perms = canRead or canWrite
            __p((perms and canRead).toString())
            __p((perms and canExecute).toString())
            val withExec = perms or canExecute
            val withoutRead = withExec and canRead.inv()
            __p((withExec).toString())
            __p((withoutRead).toString())
            __p((withoutRead and canWrite).toString())
            __p((withoutRead and canExecute).toString())
        
__check("1\n0\n7\n-8\n2\n4")
}
