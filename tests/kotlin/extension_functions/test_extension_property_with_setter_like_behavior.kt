// vybe-test: kotlin/extension_functions/test_extension_property_with_setter_like_behavior
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

class Holder(var value: Int)

        var Holder.doubled: Int
            get() = value * 2
            set(next) { value = next / 2 }

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
            val holder = Holder(3)
            holder.doubled = 10
            __p((holder.value).toString())
            __p((holder.doubled).toString())
        
__check("5\n10")
}
