// vybe-test: kotlin/kotlin_generic_variance/test_contravariant_consumer_can_downcast_target_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_generic_variance.rs

open class Animal(val kind: String)
        class Cat(value: String) : Animal(value)

        class Consumer<in T> {
            var value: T? = null
                private set

            fun push(v: T) {
                value = v
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
            val animalConsumer = Consumer<Animal>()
            val catConsumer: Consumer<Cat> = animalConsumer
            catConsumer.push(Cat("kitty"))
            __p(((catConsumer.value?.kind)).toString())
        
__check("kitty")
}
