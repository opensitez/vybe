// vybe-test: kotlin/companion_objects/test_generic_companion_factory_preserves_inferred_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder<T>(val value: T) {
            companion object {
                fun <T> make(value: T): Holder<T> = Holder(value)
            }
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
            val text = Holder.make("kotlin").value
            val number = Holder.make(12).value
            __p((text).toString())
            __p((number).toString())
        
__check("kotlin\n12")
}
