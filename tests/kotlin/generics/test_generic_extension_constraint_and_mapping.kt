// vybe-test: kotlin/generics/test_generic_extension_constraint_and_mapping
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Wrapper<T>(val value: T)

        fun <T> Wrapper<T>.map(transform: (T) -> T): T {
            return transform(this.value)
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
            __p((Wrapper("a").map { it + "b" }).toString())
            __p((Wrapper(9).map { it + 1 }).toString())
        
__check("ab\n10")
}
