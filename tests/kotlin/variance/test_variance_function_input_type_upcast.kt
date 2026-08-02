// vybe-test: kotlin/variance/test_variance_function_input_type_upcast
// origin: languages/kotlin/tests/kotlin/test_variance.rs

open class Animal
        class Cat : Animal()
        class Dog : Animal()
        fun feed(animal: Animal) { println(animal::class.simpleName) }
        fun main() {
            val cat: Cat = Cat()
            val dog: Dog = Dog()
            feed(cat)
            feed(dog)
        }

