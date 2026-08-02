// vybe-test: kotlin/kotlin_generic_variance/test_covariant_producer_can_upcast
// origin: languages/kotlin/tests/kotlin/test_kotlin_generic_variance.rs

open class Animal(val kind: String)
        class Cat(value: String) : Animal(value)
        class Dog(value: String) : Animal(value)

        class Producer<out T>(private val value: T) {
            fun value(): T = value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val cats = Producer(Cat("cat"))
            val animals: Producer<Animal> = cats
            __check((animals.value().kind).toString(), "cat")
            val dogs = Producer(Dog("dog"))
            val animalDogs: Producer<Animal> = dogs
            __check((animalDogs.value().kind).toString(), "dog")
        }
