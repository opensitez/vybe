// vybe-test: kotlin/kotlin_class_init_sequences/test_inner_constructor_for_data_properties
// origin: languages/kotlin/tests/kotlin/test_kotlin_class_init_sequences.rs

class Profile {
            val name: String
            val suffix: String
            constructor(name: String) {
                this.name = name
                this.suffix = name.takeLast(1)
            }
            constructor(name: String, idx: Int) : this(name) {
                this.suffix = name[idx]
            }

            fun render(): String = name + suffix
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
            __p((Profile("abc").render()).toString())
            __p((Profile("xyz", 1).render()).toString())
        
__check("abcc\nxyy")
}
