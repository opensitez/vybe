// vybe-test: kotlin/sealed_types/test_state_shape_preserved_in_when_mapping
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Variant {
            class A(val id: Int) : Variant()
            class B(val name: String) : Variant()
        }

        fun map(variant: Variant): String {
            return when (variant) {
                is Variant.A -> variant.id.toString()
                is Variant.B -> variant.name
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
            __p((map(Variant.A(4))).toString())
            __p((map(Variant.B("ok"))).toString())
        
__check("4\nok")
}
