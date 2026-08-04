// vybe-test: kotlin/secondary_constructors/test_constructor_super_call
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

open class Animal(val name: String)

        class Dog : Animal {
            val age: Int

            constructor(name: String, age: Int) : super(name) {
                this.age = age
            }

            constructor(name: String) : this(name, 1)
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
            val a = Dog("Rex", 5)
            val b = Dog("Buddy")
            __p((a.name).toString())
            __p((a.age).toString())
            __p((b.name).toString())
            __p((b.age).toString())
        
__check("Rex\n5\nBuddy\n1")
}
