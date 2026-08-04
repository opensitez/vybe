// vybe-test: kotlin/generics/test_generic_class_with_shadowed_type_parameter
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(val value: T) {
            fun <R> map(transform: (T) -> R): Holder<R> {
                return Holder(transform(value))
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
            val holder = Holder("7")
            val number = holder.map { it.toInt() }
            val text = holder.map { it + it }
            __p((number.value + 1).toString())
            __p((text.value).toString())
        
__check("8\n77")
}
