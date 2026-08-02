// vybe-test: kotlin/type_casts/test_unsafe_cast_to_wrong_class_is_caught
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class Animal
        class Dog : Animal()
        class Cat : Animal()

        fun main() {
            val value: Animal = Cat()
            try {
                val dog = value as Dog
                println(dog is Dog)
            } catch (e: Exception) {
                println("bad")
            }
        }

