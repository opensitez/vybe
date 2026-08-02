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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val animalConsumer = Consumer<Animal>()
            val catConsumer: Consumer<Cat> = animalConsumer
            catConsumer.push(Cat("kitty"))
            __check(((catConsumer.value?.kind)).toString(), "kitty")
        }
