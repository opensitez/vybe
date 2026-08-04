// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_super_argument_is_evaluated_for_derived_init
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

var calls = ""

        open class Parent(val seed: Int) {
            val value = seed + 1
        }

        fun computeSeed(base: Int): Int {
            calls += base.toString()
            return base * 10
        }

        class Child : Parent {
            val offset: Int

            constructor(base: Int) : super(computeSeed(base)) {
                this.offset = base + 1
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
            val c = Child(7)
            __p((c.value).toString())
            __p((c.offset).toString())
            __p((calls).toString())
        
__check("71\n8\n7")
}
