// vybe-test: kotlin/classes/test_multiple_class_instances
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Account(val id: String, var balance: Int)

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
            val a1 = Account("A", 100)
            val a2 = Account("B", 200)
            a1.balance += 50
            __p((a1.balance).toString())
            __p((a2.balance).toString())
        
__check("150\n200")
}
