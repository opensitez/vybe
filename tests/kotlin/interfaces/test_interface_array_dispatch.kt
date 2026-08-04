// vybe-test: kotlin/interfaces/test_interface_array_dispatch
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Score { fun score(): Int }
        class A : Score { override fun score(): Int = 1 }
        class B : Score { override fun score(): Int = 2 }
        class C : Score { override fun score(): Int = 3 }

        fun total(items: Array<Score>): Int {
            var sum = 0
            for (item in items) {
                sum += item.score()
            }
            return sum
        }

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
            __p((total(arrayOf(A(), B(), C()))).toString())
        
__check("6")
}
