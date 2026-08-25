// vybe-test: kotlin/kotlin_class_init_sequences/test_primary_constructor_and_init_order
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

// Damaged spelling repaired: the original opened with `class Counter start {`,
// which no Kotlin accepts — the primary-constructor parameter lost its parens
// and type. Measured under kotlinc 2.4.10: `class Counter(start: Int)` with
// the same body compiles clean and `Counter()` prints 5.
class Counter(start: Int) {
            val value: Int = start
            init {
                __p((value).toString())
            }
            constructor(): this(5)
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
            Counter()
        
__check("5")
}
